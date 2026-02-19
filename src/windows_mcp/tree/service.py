import logging
import os
import threading
import weakref
from concurrent.futures import ThreadPoolExecutor, as_completed
from time import time
from typing import TYPE_CHECKING, Any

import comtypes

from windows_mcp.native import native_capture_tree
from windows_mcp.tree.cache_utils import CachedControlHelper, CacheRequestFactory
from windows_mcp.tree.config import (
    DEFAULT_ACTIONS,
    DOCUMENT_CONTROL_TYPE_NAMES,
    INFORMATIVE_CONTROL_TYPE_NAMES,
    INTERACTIVE_CONTROL_TYPE_NAMES,
    INTERACTIVE_ROLES,
    STRUCTURAL_CONTROL_TYPE_NAMES,
    THREAD_MAX_RETRIES,
)
from windows_mcp.tree.utils import random_point_within_bounding_box
from windows_mcp.tree.views import (
    BoundingBox,
    Center,
    ScrollElementNode,
    TextElementNode,
    TreeElementNode,
    TreeState,
)
from windows_mcp.uia import (
    AccessibleRoleNames,
    Control,
    ControlFromHandle,
    PatternId,
    Rect,
    ScrollPattern,
    TreeScope,
    WindowControl,
)

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

if TYPE_CHECKING:
    from windows_mcp.desktop.service import Desktop


class Tree:
    def __init__(self, desktop: "Desktop"):
        self.desktop = weakref.proxy(desktop)
        self.screen_size = desktop.get_screen_size()
        self.screen_box = BoundingBox(
            top=0,
            left=0,
            bottom=self.screen_size.height,
            right=self.screen_size.width,
            width=self.screen_size.width,
            height=self.screen_size.height,
        )
        self.tree_state = None
        self._state_lock = threading.Lock()
        self._last_focus_event: tuple[tuple, float] | None = None

    def get_state(
        self,
        active_window_handle: int | None,
        other_windows_handles: list[int],
        use_dom: bool = False,
    ) -> TreeState:
        start_time = time()

        active_window_flag = False
        if active_window_handle:
            active_window_flag = True
            windows_handles = [active_window_handle] + other_windows_handles
        else:
            windows_handles = other_windows_handles

        interactive_nodes, scrollable_nodes, dom_informative_nodes, dom_node = (
            self.get_window_wise_nodes(
                windows_handles=windows_handles,
                active_window_flag=active_window_flag,
                use_dom=use_dom,
            )
        )
        root_node = TreeElementNode(
            name="Desktop",
            control_type="PaneControl",
            bounding_box=self.screen_box,
            center=self.screen_box.get_center(),
            window_name="Desktop",
            xpath="",
            value="",
            shortcut="",
            is_focused=False,
        )
        tree_state = TreeState(
            root_node=root_node,
            dom_node=dom_node,
            interactive_nodes=interactive_nodes,
            scrollable_nodes=scrollable_nodes,
            dom_informative_nodes=dom_informative_nodes,
        )
        with self._state_lock:
            self.tree_state = tree_state
        end_time = time()
        logger.info("Tree State capture took %.2f seconds", end_time - start_time)
        return tree_state

    # -----------------------------------------------------------------------
    # Rust fast-path: classify elements from Rust tree snapshot
    # -----------------------------------------------------------------------

    # Set of Rust control type names (without "Control" suffix) that are
    # considered interactive -- derived from INTERACTIVE_CONTROL_TYPE_NAMES.
    _RUST_INTERACTIVE_TYPES = {
        name.removesuffix("Control") for name in INTERACTIVE_CONTROL_TYPE_NAMES
    } | {name.removesuffix("Control") for name in DOCUMENT_CONTROL_TYPE_NAMES}

    # Structural types that may be scrollable containers.  Used for
    # heuristic scroll detection when live COM ScrollPattern is unavailable.
    _RUST_SCROLLABLE_CANDIDATES = {
        name.removesuffix("Control") for name in STRUCTURAL_CONTROL_TYPE_NAMES
    } | {"List", "Tree", "DataGrid", "Table"}

    def _classify_rust_tree(
        self,
        snapshot: dict,
        window_name: str,
        window_rect: list[float],
        interactive_nodes: list[TreeElementNode],
        scrollable_nodes: list[ScrollElementNode] | None = None,
    ):
        """Walk a Rust TreeElementSnapshot dict and classify elements.

        Uses only cached properties (no live COM calls).  Produces
        interactive TreeElementNodes from the Rust tree structure.

        Limitations vs the full Python path:
        - **Scroll detection** is heuristic only.  The Python path queries
          ``ScrollPattern.VerticallyScrollable`` via live COM, which requires
          a pattern invocation that cached properties cannot provide.  Here we
          detect *likely* scrollable containers by control type + child count,
          but report them with 0% scroll position.
        - **Informative text nodes** are not collected.  The Python path only
          collects these for browser DOM subtrees (``is_browser and is_dom``),
          which are never routed through the Rust fast-path.
        - **LegacyIAccessiblePattern role checks** are skipped.  Classification
          relies on ``control_type`` and ``is_keyboard_focusable`` instead.

        Args:
            snapshot: Dict from native_capture_tree (one window root).
            window_name: Corrected window name for the output nodes.
            window_rect: [left, top, right, bottom] of the window.
            interactive_nodes: Output list to append to.
            scrollable_nodes: Optional list to append heuristic scroll nodes to.
        """
        # Build window bounding box for intersection
        wl, wt, wr, wb = window_rect
        window_box_left = int(wl)
        window_box_top = int(wt)
        window_box_right = int(wr)
        window_box_bottom = int(wb)

        stack: list[dict] = [snapshot]
        while stack:
            elem = stack.pop()

            # Skip non-control / offscreen / disabled elements
            if not elem.get("is_control_element", False):
                pass  # still walk children
            elif not elem.get("is_enabled", False):
                pass
            elif elem.get("is_offscreen", False) and elem.get("control_type") != "Edit":
                pass
            else:
                control_type = elem.get("control_type", "Unknown")
                rect = elem.get("bounding_rect", [0, 0, 0, 0])
                el, et, er, eb = int(rect[0]), int(rect[1]), int(rect[2]), int(rect[3])
                area = (er - el) * (eb - et)

                if area > 0:
                    children = elem.get("children", [])

                    # Heuristic scroll detection: structural containers with
                    # children that extend beyond the container bounds are
                    # likely scrollable.  We cannot query ScrollPattern from
                    # cached data, so this is a best-effort approximation.
                    if (
                        scrollable_nodes is not None
                        and control_type in self._RUST_SCROLLABLE_CANDIDATES
                        and len(children) > 0
                    ):
                        # Check if any child extends below the container --
                        # a strong signal that the container scrolls vertically.
                        has_overflow = False
                        for child in children:
                            cr = child.get("bounding_rect", [0, 0, 0, 0])
                            if int(cr[3]) > eb:  # child bottom > parent bottom
                                has_overflow = True
                                break

                        if has_overflow:
                            il = max(window_box_left, self.screen_box.left, el)
                            it = max(window_box_top, self.screen_box.top, et)
                            ir = min(window_box_right, self.screen_box.right, er)
                            ib = min(window_box_bottom, self.screen_box.bottom, eb)
                            if ir > il and ib > it:
                                sbb = BoundingBox(
                                    left=il,
                                    top=it,
                                    right=ir,
                                    bottom=ib,
                                    width=ir - il,
                                    height=ib - it,
                                )
                                name = (
                                    elem.get("name", "").strip()
                                    or elem.get("automation_id", "")
                                    or elem.get("localized_control_type", control_type).capitalize()
                                    or "''"
                                )
                                scrollable_nodes.append(
                                    ScrollElementNode(
                                        name=name,
                                        control_type=elem.get(
                                            "localized_control_type", control_type
                                        ).title(),
                                        bounding_box=sbb,
                                        center=sbb.get_center(),
                                        xpath="",
                                        window_name=window_name,
                                        # Heuristic: assume vertical scroll, no percentage info
                                        horizontal_scrollable=False,
                                        horizontal_scroll_percent=0,
                                        vertical_scrollable=True,
                                        vertical_scroll_percent=-1,
                                        is_focused=elem.get("has_keyboard_focus", False),
                                    )
                                )

                    # Interactive classification
                    is_kb_focusable = elem.get("is_keyboard_focusable", False)
                    is_interactive = False

                    if control_type in self._RUST_INTERACTIVE_TYPES:
                        is_interactive = True
                    elif control_type == "Image" and is_kb_focusable:
                        is_interactive = True

                    if is_interactive:
                        # Compute intersection with window and screen
                        il = max(window_box_left, self.screen_box.left, el)
                        it = max(window_box_top, self.screen_box.top, et)
                        ir = min(window_box_right, self.screen_box.right, er)
                        ib = min(window_box_bottom, self.screen_box.bottom, eb)

                        if ir > il and ib > it:
                            bb = BoundingBox(
                                left=il,
                                top=it,
                                right=ir,
                                bottom=ib,
                                width=ir - il,
                                height=ib - it,
                            )
                            interactive_nodes.append(
                                TreeElementNode(
                                    name=elem.get("name", "").strip(),
                                    control_type=elem.get(
                                        "localized_control_type", control_type
                                    ).title(),
                                    bounding_box=bb,
                                    center=bb.get_center(),
                                    window_name=window_name,
                                    value="",
                                    shortcut=elem.get("accelerator_key", ""),
                                    xpath="",
                                    is_focused=elem.get("has_keyboard_focus", False),
                                )
                            )

            # Push children onto the stack (reversed for left-to-right order)
            for child in reversed(elem.get("children", [])):
                stack.append(child)  # noqa: PERF402

    def get_window_wise_nodes(
        self,
        windows_handles: list[int],
        active_window_flag: bool,
        use_dom: bool = False,
    ) -> tuple[
        list[TreeElementNode],
        list[ScrollElementNode],
        list[TextElementNode],
        ScrollElementNode | None,
    ]:
        interactive_nodes, scrollable_nodes, dom_informative_nodes = [], [], []
        dom_node: ScrollElementNode | None = None

        # Pre-calculate browser status in main thread to pass simple types to workers
        task_inputs = []
        rust_handles = []  # Non-browser windows for Rust fast-path
        for handle in windows_handles:
            is_browser = False
            try:
                # Use temporary control for property check in main thread
                # This is safe as we don't pass this specific COM object to the thread
                temp_node = ControlFromHandle(handle)
                if active_window_flag and temp_node.ClassName == "Progman":
                    continue
                is_browser = self.desktop.is_window_browser(temp_node)
            except Exception:
                pass

            if is_browser or use_dom:
                # Browser windows need full Python path (DOM mode, pattern queries)
                task_inputs.append((handle, is_browser))
            else:
                rust_handles.append(handle)

        # Rust fast-path for non-browser windows (parallel Rayon traversal)
        if rust_handles:
            rust_start = time()
            rust_tree = native_capture_tree(rust_handles, max_depth=50)
            if rust_tree is not None:
                for snapshot in rust_tree:
                    rect = snapshot.get("bounding_rect", [0, 0, 0, 0])
                    wname = snapshot.get("name", "").strip()
                    wname = self.app_name_correction(wname)
                    self._classify_rust_tree(
                        snapshot, wname, rect, interactive_nodes, scrollable_nodes
                    )
                logger.info(
                    "Rust fast-path: %d windows, %d interactive elements in %.2fs",
                    len(rust_handles),
                    len(interactive_nodes),
                    time() - rust_start,
                )
            else:
                # Rust unavailable -- fall back to Python for all
                for handle in rust_handles:
                    task_inputs.append((handle, False))

        with ThreadPoolExecutor(max_workers=min(8, os.cpu_count() or 4)) as executor:
            retry_counts = {handle: 0 for handle in windows_handles}
            future_to_handle = {
                executor.submit(self.get_nodes, handle, is_browser, use_dom): handle
                for handle, is_browser in task_inputs
            }
            while future_to_handle:  # keep running until no pending futures
                for future in as_completed(list(future_to_handle)):
                    handle = future_to_handle.pop(future)  # remove completed future
                    try:
                        result = future.result()
                        if result:
                            element_nodes, scroll_nodes, info_nodes, window_dom_node = result
                            interactive_nodes.extend(element_nodes)
                            scrollable_nodes.extend(scroll_nodes)
                            dom_informative_nodes.extend(info_nodes)
                            if window_dom_node is not None:
                                dom_node = window_dom_node
                    except Exception as e:
                        retry_counts[handle] += 1
                        logger.debug(
                            "Error in processing handle %s, retry attempt %d\nError: %s",
                            handle, retry_counts[handle], e,
                        )
                        if retry_counts[handle] < THREAD_MAX_RETRIES:
                            # Need to find is_browser again for retry
                            is_browser = next((ib for h, ib in task_inputs if h == handle), False)
                            new_future = executor.submit(
                                self.get_nodes, handle, is_browser, use_dom
                            )
                            future_to_handle[new_future] = handle
                        else:
                            logger.error(
                                "Task failed completely for handle %s after %d retries",
                                handle, THREAD_MAX_RETRIES,
                            )
        return interactive_nodes, scrollable_nodes, dom_informative_nodes, dom_node

    def iou_bounding_box(
        self,
        window_box: Rect,
        element_box: Rect,
    ) -> BoundingBox:
        # Step 1: Intersection of element and window (existing logic)
        intersection_left = max(window_box.left, element_box.left)
        intersection_top = max(window_box.top, element_box.top)
        intersection_right = min(window_box.right, element_box.right)
        intersection_bottom = min(window_box.bottom, element_box.bottom)

        # Step 2: Clamp to screen boundaries (new addition)
        intersection_left = max(self.screen_box.left, intersection_left)
        intersection_top = max(self.screen_box.top, intersection_top)
        intersection_right = min(self.screen_box.right, intersection_right)
        intersection_bottom = min(self.screen_box.bottom, intersection_bottom)

        # Step 3: Validate intersection
        if intersection_right > intersection_left and intersection_bottom > intersection_top:
            bounding_box = BoundingBox(
                left=intersection_left,
                top=intersection_top,
                right=intersection_right,
                bottom=intersection_bottom,
                width=intersection_right - intersection_left,
                height=intersection_bottom - intersection_top,
            )
        else:
            # No valid visible intersection (either outside window or screen)
            bounding_box = BoundingBox(left=0, top=0, right=0, bottom=0, width=0, height=0)
        return bounding_box

    def element_has_child_element(self, node: Control, control_type: str, child_control_type: str):
        if node.LocalizedControlType == control_type:
            first_child = node.GetFirstChildControl()
            if first_child is None:
                return False
            return first_child.LocalizedControlType == child_control_type

    def _dom_correction(
        self,
        node: Control,
        dom_interactive_nodes: list[TreeElementNode],
        window_name: str,
        dom_bounding_box: BoundingBox,
    ):
        if self.element_has_child_element(
            node, "list item", "link"
        ) or self.element_has_child_element(node, "item", "link"):
            if dom_interactive_nodes:
                dom_interactive_nodes.pop()
            return None
        elif node.ControlTypeName == "GroupControl":
            if dom_interactive_nodes:
                dom_interactive_nodes.pop()
            # Inlined is_keyboard_focusable logic for correction
            control_type_name_check = node.CachedControlTypeName
            is_kb_focusable = False
            if control_type_name_check in set(
                [
                    "EditControl",
                    "ButtonControl",
                    "CheckBoxControl",
                    "RadioButtonControl",
                    "TabItemControl",
                ]
            ):
                is_kb_focusable = True
            else:
                is_kb_focusable = node.CachedIsKeyboardFocusable

            if is_kb_focusable:
                child = node
                try:
                    while child.GetFirstChildControl() is not None:
                        if child.ControlTypeName in INTERACTIVE_CONTROL_TYPE_NAMES:
                            return None
                        child = child.GetFirstChildControl()
                except Exception:
                    return None
                if child.ControlTypeName != "TextControl":
                    return None
                legacy_pattern = node.GetLegacyIAccessiblePattern()
                value = legacy_pattern.Value
                element_bounding_box = node.BoundingRectangle
                bounding_box = self.iou_bounding_box(dom_bounding_box, element_bounding_box)
                center = bounding_box.get_center()
                is_focused = node.HasKeyboardFocus
                dom_interactive_nodes.append(
                    TreeElementNode(
                        **{
                            "name": child.Name.strip(),
                            "control_type": node.LocalizedControlType,
                            "value": value,
                            "shortcut": node.AcceleratorKey,
                            "bounding_box": bounding_box,
                            "xpath": "",
                            "center": center,
                            "window_name": window_name,
                            "is_focused": is_focused,
                        }
                    )
                )
        elif self.element_has_child_element(node, "link", "heading"):
            if dom_interactive_nodes:
                dom_interactive_nodes.pop()
            node = node.GetFirstChildControl()
            control_type = "link"
            legacy_pattern = node.GetLegacyIAccessiblePattern()
            value = legacy_pattern.Value
            element_bounding_box = node.BoundingRectangle
            bounding_box = self.iou_bounding_box(dom_bounding_box, element_bounding_box)
            center = bounding_box.get_center()
            is_focused = node.HasKeyboardFocus
            dom_interactive_nodes.append(
                TreeElementNode(
                    **{
                        "name": node.Name.strip(),
                        "control_type": control_type,
                        "value": node.Name.strip(),
                        "shortcut": node.AcceleratorKey,
                        "bounding_box": bounding_box,
                        "xpath": "",
                        "center": center,
                        "window_name": window_name,
                        "is_focused": is_focused,
                    }
                )
            )

    MAX_TREE_DEPTH = 200

    def tree_traversal(
        self,
        node: Control,
        window_bounding_box: Rect,
        window_name: str,
        is_browser: bool,
        interactive_nodes: list[TreeElementNode] | None = None,
        scrollable_nodes: list[ScrollElementNode] | None = None,
        dom_interactive_nodes: list[TreeElementNode] | None = None,
        dom_informative_nodes: list[TextElementNode] | None = None,
        is_dom: bool = False,
        is_dialog: bool = False,
        element_cache_req: Any | None = None,
        children_cache_req: Any | None = None,
        subtree_cached: bool = False,
        dom_bounding_box: BoundingBox | None = None,
        dom_context: dict | None = None,
        depth: int = 0,
    ):
        if depth >= self.MAX_TREE_DEPTH:
            logger.warning("Max tree depth %d reached, stopping traversal", self.MAX_TREE_DEPTH)
            return

        try:
            # Build cached control if per-node caching (skipped in subtree mode)
            if not subtree_cached and not hasattr(node, "_is_cached") and element_cache_req:
                node = CachedControlHelper.build_cached_control(node, element_cache_req)

            # Checks to skip the nodes that are not interactive
            is_offscreen = node.CachedIsOffscreen
            control_type_name = node.CachedControlTypeName
            # Scrollable check
            if scrollable_nodes is not None:
                if (
                    control_type_name
                    not in (INTERACTIVE_CONTROL_TYPE_NAMES | INFORMATIVE_CONTROL_TYPE_NAMES)
                ) and not is_offscreen:
                    try:
                        scroll_pattern: ScrollPattern = node.GetPattern(PatternId.ScrollPattern)
                        if scroll_pattern and scroll_pattern.VerticallyScrollable:
                            box = node.CachedBoundingRectangle
                            x, y = random_point_within_bounding_box(node=node, scale_factor=0.8)
                            center = Center(x=x, y=y)
                            name = node.CachedName
                            automation_id = node.CachedAutomationId
                            localized_control_type = node.CachedLocalizedControlType
                            has_keyboard_focus = node.CachedHasKeyboardFocus
                            scrollable_nodes.append(
                                ScrollElementNode(
                                    **{
                                        "name": name.strip()
                                        or automation_id
                                        or localized_control_type.capitalize()
                                        or "''",
                                        "control_type": localized_control_type.title(),
                                        "bounding_box": BoundingBox(
                                            **{
                                                "left": box.left,
                                                "top": box.top,
                                                "right": box.right,
                                                "bottom": box.bottom,
                                                "width": box.width(),
                                                "height": box.height(),
                                            }
                                        ),
                                        "center": center,
                                        "xpath": "",
                                        "horizontal_scrollable": scroll_pattern.HorizontallyScrollable,
                                        "horizontal_scroll_percent": scroll_pattern.HorizontalScrollPercent
                                        if scroll_pattern.HorizontallyScrollable
                                        else 0,
                                        "vertical_scrollable": scroll_pattern.VerticallyScrollable,
                                        "vertical_scroll_percent": scroll_pattern.VerticalScrollPercent
                                        if scroll_pattern.VerticallyScrollable
                                        else 0,
                                        "window_name": window_name,
                                        "is_focused": has_keyboard_focus,
                                    }
                                )
                            )
                    except Exception:
                        logger.debug("Scroll pattern query failed for node", exc_info=True)

            # Interactive and Informative checks
            # Pre-calculate common properties
            is_control_element = node.CachedIsControlElement
            element_bounding_box = node.CachedBoundingRectangle
            width = element_bounding_box.width()
            height = element_bounding_box.height()
            area = width * height

            # Is Visible Check
            is_visible = (
                (area > 0)
                and (not is_offscreen or control_type_name == "EditControl")
                and is_control_element
            )

            if is_visible:
                is_enabled = node.CachedIsEnabled
                if is_enabled:
                    # Determine is_keyboard_focusable
                    if control_type_name in set(
                        [
                            "EditControl",
                            "ButtonControl",
                            "CheckBoxControl",
                            "RadioButtonControl",
                            "TabItemControl",
                        ]
                    ):
                        is_keyboard_focusable = True
                    else:
                        is_keyboard_focusable = node.CachedIsKeyboardFocusable

                    # Interactive Check
                    if interactive_nodes is not None:
                        is_interactive = False
                        # Cache LegacyIAccessiblePattern per element (was 3-4 COM calls)
                        legacy_pattern = None

                        if (
                            is_browser
                            and control_type_name in set(["DataItemControl", "ListItemControl"])
                            and not is_keyboard_focusable
                        ):
                            is_interactive = False
                        elif (
                            not is_browser
                            and control_type_name == "ImageControl"
                            and is_keyboard_focusable
                        ):
                            is_interactive = True
                        elif control_type_name in (
                            INTERACTIVE_CONTROL_TYPE_NAMES | DOCUMENT_CONTROL_TYPE_NAMES
                        ):
                            # Role check
                            try:
                                if legacy_pattern is None:
                                    legacy_pattern = node.GetLegacyIAccessiblePattern()
                                is_role_interactive = (
                                    AccessibleRoleNames.get(legacy_pattern.Role, "Default")
                                    in INTERACTIVE_ROLES
                                )
                            except Exception:
                                is_role_interactive = False

                            # Image check
                            is_image = False
                            if control_type_name == "ImageControl":  # approximated
                                localized = node.CachedLocalizedControlType
                                if localized == "graphic" or not is_keyboard_focusable:
                                    is_image = True

                            if is_role_interactive and (not is_image or is_keyboard_focusable):
                                is_interactive = True

                        elif control_type_name == "GroupControl":
                            if is_browser:
                                try:
                                    if legacy_pattern is None:
                                        legacy_pattern = node.GetLegacyIAccessiblePattern()
                                    is_role_interactive = (
                                        AccessibleRoleNames.get(legacy_pattern.Role, "Default")
                                        in INTERACTIVE_ROLES
                                    )
                                except Exception:
                                    is_role_interactive = False

                                is_default_action = False
                                try:
                                    if legacy_pattern is None:
                                        legacy_pattern = node.GetLegacyIAccessiblePattern()
                                    if legacy_pattern.DefaultAction.title() in DEFAULT_ACTIONS:
                                        is_default_action = True
                                except Exception:
                                    pass

                                if is_role_interactive and (
                                    is_default_action or is_keyboard_focusable
                                ):
                                    is_interactive = True

                        if is_interactive:
                            if legacy_pattern is None:
                                legacy_pattern = node.GetLegacyIAccessiblePattern()
                            value = (
                                legacy_pattern.Value.strip()
                                if legacy_pattern.Value is not None
                                else ""
                            )
                            is_focused = node.CachedHasKeyboardFocus
                            name = node.CachedName.strip()
                            localized_control_type = node.CachedLocalizedControlType
                            accelerator_key = node.CachedAcceleratorKey

                            if is_browser and is_dom:
                                bounding_box = self.iou_bounding_box(
                                    dom_bounding_box, element_bounding_box
                                )
                                center = bounding_box.get_center()
                                tree_node = TreeElementNode(
                                    **{
                                        "name": name,
                                        "control_type": localized_control_type.title(),
                                        "value": value,
                                        "shortcut": accelerator_key,
                                        "bounding_box": bounding_box,
                                        "center": center,
                                        "xpath": "",
                                        "window_name": window_name,
                                        "is_focused": is_focused,
                                    }
                                )
                                dom_interactive_nodes.append(tree_node)
                                self._dom_correction(
                                    node, dom_interactive_nodes, window_name, dom_bounding_box
                                )
                            else:
                                bounding_box = self.iou_bounding_box(
                                    window_bounding_box, element_bounding_box
                                )
                                center = bounding_box.get_center()
                                tree_node = TreeElementNode(
                                    **{
                                        "name": name,
                                        "control_type": localized_control_type.title(),
                                        "value": value,
                                        "shortcut": accelerator_key,
                                        "bounding_box": bounding_box,
                                        "center": center,
                                        "xpath": "",
                                        "window_name": window_name,
                                        "is_focused": is_focused,
                                    }
                                )
                                interactive_nodes.append(tree_node)

                    # Informative Check
                    if dom_informative_nodes is not None:
                        # is_element_text check
                        is_text = False
                        if control_type_name in INFORMATIVE_CONTROL_TYPE_NAMES:
                            # is_element_image: True when ImageControl is not focusable OR
                            # its localized type is "graphic"
                            is_image_check = False
                            if control_type_name == "ImageControl":
                                localized = node.CachedLocalizedControlType
                                if not is_keyboard_focusable or localized == "graphic":
                                    is_image_check = True

                            if not is_image_check:
                                is_text = True

                        if is_text:
                            if is_browser and is_dom:
                                name = node.CachedName
                                dom_informative_nodes.append(
                                    TextElementNode(
                                        text=name.strip(),
                                    )
                                )

            # Phase 3: Children Retrieval
            if subtree_cached:
                # Subtree pre-cached: walk cached tree directly (no COM calls)
                try:
                    children = node.GetCachedChildren()
                    for child in children:
                        child._is_cached = True
                except Exception:
                    children = CachedControlHelper.get_cached_children(node, children_cache_req)
            else:
                children = CachedControlHelper.get_cached_children(node, children_cache_req)

            # Recursively traverse the tree the right to left for normal apps and for DOM traverse from left to right
            for child in children if is_dom else children[::-1]:
                # Incrementally building the xpath

                # Check if the child is a DOM element
                if is_browser and child.CachedAutomationId == "RootWebArea":
                    rect = child.CachedBoundingRectangle
                    child_dom_bb = BoundingBox(
                        left=rect.left,
                        top=rect.top,
                        right=rect.right,
                        bottom=rect.bottom,
                        width=rect.width(),
                        height=rect.height(),
                    )
                    if dom_context is not None:
                        dom_context["dom"] = child
                        dom_context["dom_bounding_box"] = child_dom_bb
                    # enter DOM subtree
                    self.tree_traversal(
                        child,
                        window_bounding_box,
                        window_name,
                        is_browser,
                        interactive_nodes,
                        scrollable_nodes,
                        dom_interactive_nodes,
                        dom_informative_nodes,
                        is_dom=True,
                        is_dialog=is_dialog,
                        element_cache_req=element_cache_req,
                        children_cache_req=children_cache_req,
                        subtree_cached=subtree_cached,
                        dom_bounding_box=child_dom_bb,
                        dom_context=dom_context,
                        depth=depth + 1,
                    )
                # Check if the child is a dialog
                elif isinstance(child, WindowControl):
                    if not child.CachedIsOffscreen:
                        if is_dom and dom_bounding_box:
                            rect = child.CachedBoundingRectangle
                            if rect.width() > 0.8 * dom_bounding_box.width:
                                # Because this window element covers the majority of the screen
                                dom_interactive_nodes.clear()
                        else:
                            # Inline is_window_modal
                            is_modal = False
                            try:
                                window_pattern = child.GetWindowPattern()
                                is_modal = window_pattern.IsModal
                            except Exception:
                                pass

                            if is_modal:
                                # Because this window element is modal
                                interactive_nodes.clear()
                    # enter dialog subtree
                    self.tree_traversal(
                        child,
                        window_bounding_box,
                        window_name,
                        is_browser,
                        interactive_nodes,
                        scrollable_nodes,
                        dom_interactive_nodes,
                        dom_informative_nodes,
                        is_dom=is_dom,
                        is_dialog=True,
                        element_cache_req=element_cache_req,
                        children_cache_req=children_cache_req,
                        subtree_cached=subtree_cached,
                        dom_bounding_box=dom_bounding_box,
                        dom_context=dom_context,
                        depth=depth + 1,
                    )
                else:
                    # normal non-dialog children
                    self.tree_traversal(
                        child,
                        window_bounding_box,
                        window_name,
                        is_browser,
                        interactive_nodes,
                        scrollable_nodes,
                        dom_interactive_nodes,
                        dom_informative_nodes,
                        is_dom=is_dom,
                        is_dialog=is_dialog,
                        element_cache_req=element_cache_req,
                        children_cache_req=children_cache_req,
                        subtree_cached=subtree_cached,
                        dom_bounding_box=dom_bounding_box,
                        dom_context=dom_context,
                        depth=depth + 1,
                    )
        except Exception as e:
            logger.error("Error in tree_traversal: %s", e, exc_info=True)
            raise

    def app_name_correction(self, app_name: str) -> str:
        match app_name:
            case "Progman":
                return "Desktop"
            case "Shell_TrayWnd" | "Shell_SecondaryTrayWnd":
                return "Taskbar"
            case "Microsoft.UI.Content.PopupWindowSiteBridge":
                return "Context Menu"
            case _:
                return app_name

    def get_nodes(
        self, handle: int, is_browser: bool = False, use_dom: bool = False
    ) -> tuple[
        list[TreeElementNode],
        list[ScrollElementNode],
        list[TextElementNode],
        ScrollElementNode | None,
    ]:
        com_initialized = False
        try:
            comtypes.CoInitialize()
            com_initialized = True
            # Rehydrate Control from handle within the thread's COM context
            node = ControlFromHandle(handle)
            if not node:
                raise Exception("Failed to create Control from handle")

            # Try single-call subtree caching (1 COM call vs 2N per-node calls)
            subtree_cached = False
            try:
                subtree_req = CacheRequestFactory.create_subtree_cache()
                cached_root = node.BuildUpdatedCache(subtree_req)
                cached_root._is_cached = True
                node = cached_root
                subtree_cached = True
            except Exception as e:
                logger.debug("Subtree caching failed for handle %s: %s", handle, e)

            # Fallback: per-node cache requests
            element_cache_req = None
            children_cache_req = None
            if not subtree_cached:
                element_cache_req = CacheRequestFactory.create_tree_traversal_cache()
                element_cache_req.TreeScope = TreeScope.TreeScope_Element
                children_cache_req = CacheRequestFactory.create_tree_traversal_cache()
                children_cache_req.TreeScope = (
                    TreeScope.TreeScope_Element | TreeScope.TreeScope_Children
                )

            if subtree_cached:
                window_bounding_box = node.CachedBoundingRectangle
            else:
                window_bounding_box = node.BoundingRectangle

            (
                interactive_nodes,
                dom_interactive_nodes,
                dom_informative_nodes,
                scrollable_nodes,
            ) = [], [], [], []
            if subtree_cached:
                window_name = node.CachedName.strip()
            else:
                window_name = node.Name.strip()
            window_name = self.app_name_correction(window_name)

            # Thread-local DOM context (avoids cross-thread COM object sharing)
            dom_context: dict = {}

            self.tree_traversal(
                node,
                window_bounding_box,
                window_name,
                is_browser,
                interactive_nodes,
                scrollable_nodes,
                dom_interactive_nodes,
                dom_informative_nodes,
                is_dom=False,
                is_dialog=False,
                element_cache_req=element_cache_req,
                children_cache_req=children_cache_req,
                subtree_cached=subtree_cached,
                dom_context=dom_context,
            )
            logger.debug("Window name:%s", window_name)
            logger.debug("Interactive nodes:%d", len(interactive_nodes))
            if is_browser:
                logger.debug("DOM interactive nodes:%d", len(dom_interactive_nodes))
                logger.debug("DOM informative nodes:%d", len(dom_informative_nodes))
            logger.debug("Scrollable nodes:%d", len(scrollable_nodes))

            # Extract DOM scroll info in worker thread (COM object stays in its apartment)
            dom_node: ScrollElementNode | None = None
            if dom_context.get("dom"):
                dom_element = dom_context["dom"]
                dom_bb = dom_context["dom_bounding_box"]
                try:
                    scroll_pattern: ScrollPattern = dom_element.GetPattern(PatternId.ScrollPattern)
                    dom_node = ScrollElementNode(
                        name="DOM",
                        control_type="DocumentControl",
                        bounding_box=dom_bb,
                        center=dom_bb.get_center(),
                        horizontal_scrollable=scroll_pattern.HorizontallyScrollable
                        if scroll_pattern
                        else False,
                        horizontal_scroll_percent=scroll_pattern.HorizontalScrollPercent
                        if scroll_pattern and scroll_pattern.HorizontallyScrollable
                        else 0,
                        vertical_scrollable=scroll_pattern.VerticallyScrollable
                        if scroll_pattern
                        else False,
                        vertical_scroll_percent=scroll_pattern.VerticalScrollPercent
                        if scroll_pattern and scroll_pattern.VerticallyScrollable
                        else 0,
                        xpath="",
                        window_name="DOM",
                        is_focused=False,
                    )
                except Exception as e:
                    logger.debug("Failed to extract DOM scroll info: %s", e)

            if use_dom:
                if is_browser:
                    return (
                        dom_interactive_nodes,
                        scrollable_nodes,
                        dom_informative_nodes,
                        dom_node,
                    )
                else:
                    return ([], [], [], None)
            else:
                interactive_nodes.extend(dom_interactive_nodes)
                return (interactive_nodes, scrollable_nodes, dom_informative_nodes, dom_node)
        except Exception as e:
            logger.error("Error getting nodes for handle %s: %s", handle, e)
            raise
        finally:
            if com_initialized:
                comtypes.CoUninitialize()

    def _on_focus_change(self, sender: Any):
        """Handle focus change events."""
        # Debounce duplicate events
        current_time = time()
        element = Control.CreateControlFromElement(sender)
        runtime_id = element.GetRuntimeId()
        event_key = tuple(runtime_id)
        if self._last_focus_event is not None:
            last_key, last_time = self._last_focus_event
            if last_key == event_key and (current_time - last_time) < 1.0:
                return None
        self._last_focus_event = (event_key, current_time)

        try:
            logger.debug(
                "[WatchDog] Focus changed to: '%s' (%s)",
                element.Name, element.ControlTypeName,
            )
        except Exception:
            pass

    def _on_property_change(self, sender: Any, propertyId: int, newValue):
        """Handle property change events."""
        try:
            element = Control.CreateControlFromElement(sender)
            logger.debug(
                "[WatchDog] Property changed: ID=%s Value=%s Element: '%s' (%s)",
                propertyId, newValue, element.Name, element.ControlTypeName,
            )
        except Exception:
            pass
