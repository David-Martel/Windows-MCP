//! UIA accessibility tree traversal via `windows-rs` and Rayon.
//!
//! [`capture_tree_raw`] captures the accessibility tree for one or more
//! windows using a single `BuildUpdatedCache(TreeScope_Subtree)` call per
//! window, parallelised across Rayon worker threads.
//!
//! # COM apartment model
//!
//! Each Rayon thread initialises its own MTA COM apartment via `COMGuard`.
//! COM interfaces are never shared across thread boundaries.

pub mod element;

use element::TreeElementSnapshot;

use rayon::prelude::*;
use windows::Win32::System::Com::{CoCreateInstance, CLSCTX_INPROC_SERVER};
use windows::Win32::UI::Accessibility::{
    CUIAutomation, IUIAutomation, IUIAutomationCacheRequest, IUIAutomationElement,
    IUIAutomationElementArray, TreeScope_Subtree, UIA_AcceleratorKeyPropertyId,
    UIA_AutomationIdPropertyId, UIA_BoundingRectanglePropertyId, UIA_ClassNamePropertyId,
    UIA_ControlTypePropertyId, UIA_HasKeyboardFocusPropertyId, UIA_IsControlElementPropertyId,
    UIA_IsEnabledPropertyId, UIA_IsKeyboardFocusablePropertyId, UIA_IsOffscreenPropertyId,
    UIA_LocalizedControlTypePropertyId, UIA_NamePropertyId,
    UIA_AppBarControlTypeId, UIA_ButtonControlTypeId, UIA_CalendarControlTypeId,
    UIA_CheckBoxControlTypeId, UIA_ComboBoxControlTypeId, UIA_CustomControlTypeId,
    UIA_DataGridControlTypeId, UIA_DataItemControlTypeId, UIA_DocumentControlTypeId,
    UIA_EditControlTypeId, UIA_GroupControlTypeId, UIA_HeaderControlTypeId,
    UIA_HeaderItemControlTypeId, UIA_HyperlinkControlTypeId, UIA_ImageControlTypeId,
    UIA_ListControlTypeId, UIA_ListItemControlTypeId, UIA_MenuBarControlTypeId,
    UIA_MenuControlTypeId, UIA_MenuItemControlTypeId, UIA_PaneControlTypeId,
    UIA_ProgressBarControlTypeId, UIA_RadioButtonControlTypeId, UIA_ScrollBarControlTypeId,
    UIA_SemanticZoomControlTypeId, UIA_SeparatorControlTypeId, UIA_SliderControlTypeId,
    UIA_SpinnerControlTypeId, UIA_SplitButtonControlTypeId, UIA_StatusBarControlTypeId,
    UIA_TabControlTypeId, UIA_TabItemControlTypeId, UIA_TableControlTypeId,
    UIA_TextControlTypeId, UIA_ThumbControlTypeId, UIA_TitleBarControlTypeId,
    UIA_ToolBarControlTypeId, UIA_ToolTipControlTypeId, UIA_TreeControlTypeId,
    UIA_TreeItemControlTypeId, UIA_WindowControlTypeId, UIA_CONTROLTYPE_ID,
};
use windows::Win32::Foundation::HWND;

use crate::com::COMGuard;
use crate::errors::WindowsMcpError;

// ---------------------------------------------------------------------------
// Control-type ID -> name mapping
// ---------------------------------------------------------------------------

fn control_type_name(id: UIA_CONTROLTYPE_ID) -> &'static str {
    match id {
        x if x == UIA_AppBarControlTypeId => "AppBar",
        x if x == UIA_ButtonControlTypeId => "Button",
        x if x == UIA_CalendarControlTypeId => "Calendar",
        x if x == UIA_CheckBoxControlTypeId => "CheckBox",
        x if x == UIA_ComboBoxControlTypeId => "ComboBox",
        x if x == UIA_CustomControlTypeId => "Custom",
        x if x == UIA_DataGridControlTypeId => "DataGrid",
        x if x == UIA_DataItemControlTypeId => "DataItem",
        x if x == UIA_DocumentControlTypeId => "Document",
        x if x == UIA_EditControlTypeId => "Edit",
        x if x == UIA_GroupControlTypeId => "Group",
        x if x == UIA_HeaderControlTypeId => "Header",
        x if x == UIA_HeaderItemControlTypeId => "HeaderItem",
        x if x == UIA_HyperlinkControlTypeId => "Hyperlink",
        x if x == UIA_ImageControlTypeId => "Image",
        x if x == UIA_ListControlTypeId => "List",
        x if x == UIA_ListItemControlTypeId => "ListItem",
        x if x == UIA_MenuBarControlTypeId => "MenuBar",
        x if x == UIA_MenuControlTypeId => "Menu",
        x if x == UIA_MenuItemControlTypeId => "MenuItem",
        x if x == UIA_PaneControlTypeId => "Pane",
        x if x == UIA_ProgressBarControlTypeId => "ProgressBar",
        x if x == UIA_RadioButtonControlTypeId => "RadioButton",
        x if x == UIA_ScrollBarControlTypeId => "ScrollBar",
        x if x == UIA_SemanticZoomControlTypeId => "SemanticZoom",
        x if x == UIA_SeparatorControlTypeId => "Separator",
        x if x == UIA_SliderControlTypeId => "Slider",
        x if x == UIA_SpinnerControlTypeId => "Spinner",
        x if x == UIA_SplitButtonControlTypeId => "SplitButton",
        x if x == UIA_StatusBarControlTypeId => "StatusBar",
        x if x == UIA_TabControlTypeId => "Tab",
        x if x == UIA_TabItemControlTypeId => "TabItem",
        x if x == UIA_TableControlTypeId => "Table",
        x if x == UIA_TextControlTypeId => "Text",
        x if x == UIA_ThumbControlTypeId => "Thumb",
        x if x == UIA_TitleBarControlTypeId => "TitleBar",
        x if x == UIA_ToolBarControlTypeId => "ToolBar",
        x if x == UIA_ToolTipControlTypeId => "ToolTip",
        x if x == UIA_TreeControlTypeId => "Tree",
        x if x == UIA_TreeItemControlTypeId => "TreeItem",
        x if x == UIA_WindowControlTypeId => "Window",
        _ => "Unknown",
    }
}

// ---------------------------------------------------------------------------
// Cache request builder
// ---------------------------------------------------------------------------

unsafe fn build_cache_request(
    uia: &IUIAutomation,
) -> Result<IUIAutomationCacheRequest, WindowsMcpError> {
    let req = uia
        .CreateCacheRequest()
        .map_err(|e| WindowsMcpError::ComError(format!("CreateCacheRequest: {e}")))?;

    req.SetTreeScope(TreeScope_Subtree)
        .map_err(|e| WindowsMcpError::ComError(format!("SetTreeScope: {e}")))?;

    let properties = [
        UIA_NamePropertyId,
        UIA_AutomationIdPropertyId,
        UIA_ControlTypePropertyId,
        UIA_LocalizedControlTypePropertyId,
        UIA_ClassNamePropertyId,
        UIA_BoundingRectanglePropertyId,
        UIA_IsOffscreenPropertyId,
        UIA_IsEnabledPropertyId,
        UIA_IsControlElementPropertyId,
        UIA_HasKeyboardFocusPropertyId,
        UIA_IsKeyboardFocusablePropertyId,
        UIA_AcceleratorKeyPropertyId,
    ];
    for prop in properties {
        req.AddProperty(prop)
            .map_err(|e| WindowsMcpError::ComError(format!("AddProperty({prop:?}): {e}")))?;
    }

    Ok(req)
}

// ---------------------------------------------------------------------------
// Recursive tree walker
// ---------------------------------------------------------------------------

macro_rules! bstr_or_empty {
    ($expr:expr) => {
        unsafe { $expr }
            .map(|b: windows::core::BSTR| b.to_string())
            .unwrap_or_default()
    };
}

macro_rules! bool_or_false {
    ($expr:expr) => {
        unsafe { $expr }
            .map(|b: windows::Win32::Foundation::BOOL| b.as_bool())
            .unwrap_or(false)
    };
}

unsafe fn walk_element(
    element: &IUIAutomationElement,
    depth: usize,
    max_depth: usize,
) -> TreeElementSnapshot {
    let name = bstr_or_empty!(element.CachedName());
    let automation_id = bstr_or_empty!(element.CachedAutomationId());
    let localized_control_type = bstr_or_empty!(element.CachedLocalizedControlType());
    let class_name = bstr_or_empty!(element.CachedClassName());
    let accelerator_key = bstr_or_empty!(element.CachedAcceleratorKey());

    let control_type = element
        .CachedControlType()
        .map(|id| control_type_name(id).to_owned())
        .unwrap_or_else(|_| "Unknown".to_owned());

    let bounding_rect = element
        .CachedBoundingRectangle()
        .map(|r| [r.left as f64, r.top as f64, r.right as f64, r.bottom as f64])
        .unwrap_or([0.0, 0.0, 0.0, 0.0]);

    let is_offscreen = bool_or_false!(element.CachedIsOffscreen());
    let is_enabled = bool_or_false!(element.CachedIsEnabled());
    let is_control_element = bool_or_false!(element.CachedIsControlElement());
    let has_keyboard_focus = bool_or_false!(element.CachedHasKeyboardFocus());
    let is_keyboard_focusable = bool_or_false!(element.CachedIsKeyboardFocusable());

    let children = if depth < max_depth {
        collect_children(element, depth, max_depth)
    } else {
        Vec::new()
    };

    TreeElementSnapshot {
        name,
        automation_id,
        control_type,
        localized_control_type,
        class_name,
        bounding_rect,
        is_offscreen,
        is_enabled,
        is_control_element,
        has_keyboard_focus,
        is_keyboard_focusable,
        accelerator_key,
        depth,
        children,
    }
}

/// Maximum children per node to prevent memory exhaustion on
/// pathological trees (e.g. a grid with 100k cells).
const MAX_CHILDREN_PER_NODE: i32 = 512;

unsafe fn collect_children(
    parent: &IUIAutomationElement,
    depth: usize,
    max_depth: usize,
) -> Vec<TreeElementSnapshot> {
    let array: IUIAutomationElementArray = match parent.GetCachedChildren() {
        Ok(arr) => arr,
        Err(_) => return Vec::new(),
    };

    let len = match array.Length() {
        Ok(n) if n > 0 => n.min(MAX_CHILDREN_PER_NODE),
        _ => return Vec::new(),
    };

    let mut children = Vec::with_capacity(len as usize);
    for i in 0..len {
        if let Ok(child) = array.GetElement(i) {
            children.push(walk_element(&child, depth + 1, max_depth));
        }
    }
    children
}

// ---------------------------------------------------------------------------
// Per-window traversal (runs inside a Rayon task)
// ---------------------------------------------------------------------------

fn capture_window(handle: isize, max_depth: usize) -> Option<TreeElementSnapshot> {
    let _com_guard = COMGuard::init().ok()?;

    let uia: IUIAutomation =
        unsafe { CoCreateInstance(&CUIAutomation, None, CLSCTX_INPROC_SERVER).ok()? };

    let cache_req = unsafe { build_cache_request(&uia).ok()? };

    let root: IUIAutomationElement = unsafe {
        uia.ElementFromHandleBuildCache(HWND(handle as *mut core::ffi::c_void), &cache_req)
            .ok()?
    };

    let snapshot = unsafe { walk_element(&root, 0, max_depth) };
    Some(snapshot)
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/// Capture the Windows UI Automation accessibility tree for one or more
/// windows, returning owned snapshots.
///
/// Windows are traversed in parallel using Rayon.  Each thread initialises
/// its own COM apartment.  Invalid/inaccessible handles are silently skipped.
///
/// `max_depth` is clamped to 50 to stay within Rayon's ~2MB thread stack.
/// Each recursion level uses ~1-2 KB of stack, so 50 levels ≈ 50-100 KB.
pub fn capture_tree_raw(window_handles: &[isize], max_depth: usize) -> Vec<TreeElementSnapshot> {
    let max_depth = max_depth.min(50);

    window_handles
        .par_iter()
        .copied()
        .filter(|&handle| handle != 0)
        .filter_map(|handle| capture_window(handle, max_depth))
        .collect()
}
