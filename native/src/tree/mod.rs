//! UIA accessibility tree traversal via `windows-rs` and Rayon.
//!
//! # Overview
//!
//! [`capture_tree`] is a `#[pyfunction]` that:
//!
//! 1. Releases the GIL with `py.allow_threads()`.
//! 2. Spawns one Rayon task per window handle using `rayon::iter::par_iter`.
//! 3. Each task initialises its own COM MTA apartment (`CoInitializeEx`),
//!    creates a fresh `IUIAutomation` instance, attaches a `CacheRequest`
//!    covering the whole subtree, and walks the cached tree recursively.
//! 4. The resulting [`TreeElementSnapshot`] trees are sent back to the main
//!    thread and converted to Python `dict` objects after the GIL is
//!    re-acquired.
//!
//! # COM apartment model
//!
//! Windows-MCP's main thread uses STA (Single-Threaded Apartment) for
//! compatibility with the existing comtypes-based Python code.  Rayon worker
//! threads are MTA (Multi-Threaded Apartment) which is acceptable here
//! because we create a **new** `IUIAutomation` instance per thread -- we
//! never share COM pointers across apartment boundaries.
//!
//! Each Rayon thread receives `COINIT_MULTITHREADED`.  `S_FALSE` from
//! `CoInitializeEx` means the thread already has an MTA apartment
//! (possible if Rayon reuses a thread) -- that is not an error.
//!
//! # Performance
//!
//! The single `BuildUpdatedCache` call with `TreeScope_Subtree` batches all
//! cross-process COM round-trips into one RPC.  This is 10-100x faster than
//! calling `BuildUpdatedCache` per-node (the Python comtypes approach).
//!
//! # Safety invariants
//!
//! - COM interfaces (`IUIAutomation`, `IUIAutomationElement`, …) are **not**
//!   `Send`.  All COM work is confined to the lambda passed to
//!   `py.allow_threads()` -- the lambda moves owned Rust data out and returns
//!   plain Rust structs (`Vec<Result<TreeElementSnapshot, …>>`).
//! - `COMGuard` calls `CoUninitialize` on `Drop`, ensuring the COM apartment
//!   is cleaned up even on early return or panic.

pub mod element;

use element::TreeElementSnapshot;

use pyo3::prelude::*;
use pyo3::types::PyList;
use rayon::prelude::*;
use windows::Win32::System::Com::{
    CoCreateInstance, CoInitializeEx, CoUninitialize, CLSCTX_ALL, COINIT_MULTITHREADED,
};
use windows::Win32::UI::Accessibility::{
    CUIAutomation, IUIAutomation, IUIAutomationCacheRequest, IUIAutomationElement,
    IUIAutomationElementArray, TreeScope_Subtree, UIA_AcceleratorKeyPropertyId,
    UIA_AutomationIdPropertyId, UIA_BoundingRectanglePropertyId, UIA_ClassNamePropertyId,
    UIA_ControlTypePropertyId, UIA_HasKeyboardFocusPropertyId, UIA_IsControlElementPropertyId,
    UIA_IsEnabledPropertyId, UIA_IsKeyboardFocusablePropertyId, UIA_IsOffscreenPropertyId,
    UIA_LocalizedControlTypePropertyId, UIA_NamePropertyId,
    // Control-type ID constants for the name lookup.
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

use crate::errors::WindowsMcpError;

// ---------------------------------------------------------------------------
// COM apartment RAII guard
// ---------------------------------------------------------------------------

/// RAII wrapper that calls [`CoUninitialize`] on [`Drop`] when appropriate.
///
/// Instantiate **once per thread** via [`COMGuard::init`].  The guard tracks
/// whether `CoInitializeEx` actually succeeded (vs. `RPC_E_CHANGED_MODE`)
/// and only calls `CoUninitialize` when a balancing call is required per MSDN.
///
/// The `PhantomData<*const ()>` field enforces `!Send` + `!Sync` at compile
/// time, preventing the guard from being moved across thread boundaries.
///
/// ```rust,ignore
/// let _com = COMGuard::init()?;
/// // ... COM work ...
/// // CoUninitialize called automatically here (only if init succeeded)
/// ```
struct COMGuard {
    /// Whether `CoUninitialize` should be called on drop.
    /// `true` for `S_OK` and `S_FALSE`; `false` for `RPC_E_CHANGED_MODE`.
    should_uninit: bool,
    /// Enforce `!Send` + `!Sync` -- COM apartments are per-thread.
    _not_send: std::marker::PhantomData<*const ()>,
}

impl COMGuard {
    /// Initialise (or join) the thread's MTA COM apartment.
    ///
    /// Returns `Ok(COMGuard)` for `S_OK` (newly initialised), `S_FALSE`
    /// (already initialised -- `CoUninitialize` still required to balance),
    /// and `RPC_E_CHANGED_MODE` (thread has STA; COM is usable but we must
    /// NOT call `CoUninitialize` since we did not successfully initialise).
    fn init() -> Result<Self, WindowsMcpError> {
        // SAFETY: pvReserved must be NULL per MSDN -- we pass None.
        let hr = unsafe { CoInitializeEx(None, COINIT_MULTITHREADED) };

        let hresult_value = hr.0 as u32;
        match hresult_value {
            // S_OK -- newly initialised, must call CoUninitialize.
            0x0 => Ok(COMGuard {
                should_uninit: true,
                _not_send: std::marker::PhantomData,
            }),
            // S_FALSE -- already initialised on this thread, must still
            // call CoUninitialize to balance the reference count.
            0x1 => Ok(COMGuard {
                should_uninit: true,
                _not_send: std::marker::PhantomData,
            }),
            // RPC_E_CHANGED_MODE (0x80010106) -- thread already has STA.
            // COM is usable in STA mode, but we must NOT call CoUninitialize
            // because this call did not succeed per MSDN.
            0x80010106 => Ok(COMGuard {
                should_uninit: false,
                _not_send: std::marker::PhantomData,
            }),
            _ => Err(WindowsMcpError::ComError(format!(
                "CoInitializeEx failed: HRESULT 0x{hresult_value:08X}"
            ))),
        }
    }
}

impl Drop for COMGuard {
    fn drop(&mut self) {
        if self.should_uninit {
            // SAFETY: balances a successful CoInitializeEx call (S_OK or S_FALSE).
            unsafe { CoUninitialize() };
        }
    }
}

// ---------------------------------------------------------------------------
// Control-type ID → name mapping
// ---------------------------------------------------------------------------

/// Translate a raw `UIA_CONTROLTYPE_ID` integer to a human-readable string.
///
/// Returns `"Unknown"` for any ID not in the table rather than panicking.
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

/// Build a `IUIAutomationCacheRequest` that covers the full subtree and
/// pre-fetches every property we need in a single cross-process RPC.
///
/// # Safety
///
/// `uia` must be a valid `IUIAutomation` obtained on the calling thread's COM
/// apartment.  The returned `IUIAutomationCacheRequest` is only valid on the
/// same thread.
unsafe fn build_cache_request(
    uia: &IUIAutomation,
) -> Result<IUIAutomationCacheRequest, WindowsMcpError> {
    let req = uia
        .CreateCacheRequest()
        .map_err(|e| WindowsMcpError::ComError(format!("CreateCacheRequest: {e}")))?;

    // Scope: capture element itself and all descendants in one RPC call.
    req.SetTreeScope(TreeScope_Subtree)
        .map_err(|e| WindowsMcpError::ComError(format!("SetTreeScope: {e}")))?;

    // Pre-fetch each property we will read from the cache.
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

/// Extract a `BSTR` cached property as a UTF-8 `String`, returning an empty
/// string on any error (property unavailable / not cached).
///
/// # Safety
///
/// `element` must be a valid cached `IUIAutomationElement` on the calling
/// thread.
macro_rules! bstr_or_empty {
    ($expr:expr) => {
        unsafe { $expr }
            .map(|b: windows::core::BSTR| b.to_string())
            .unwrap_or_default()
    };
}

/// Extract a `BOOL` cached property as a Rust `bool`, defaulting to `false`.
macro_rules! bool_or_false {
    ($expr:expr) => {
        unsafe { $expr }
            .map(|b: windows::Win32::Foundation::BOOL| b.as_bool())
            .unwrap_or(false)
    };
}

/// Recursively walk a cached UIA element tree and collect snapshots.
///
/// `element` must already have been populated by a `BuildUpdatedCache` call
/// with `TreeScope_Subtree` so that all child links and properties are
/// available in the local cache without further COM round-trips.
///
/// # Safety
///
/// `element` must be a valid `IUIAutomationElement` on the calling thread's
/// COM apartment.
unsafe fn walk_element(
    element: &IUIAutomationElement,
    depth: usize,
    max_depth: usize,
) -> TreeElementSnapshot {
    // Read all cached string properties.
    let name = bstr_or_empty!(element.CachedName());
    let automation_id = bstr_or_empty!(element.CachedAutomationId());
    let localized_control_type = bstr_or_empty!(element.CachedLocalizedControlType());
    let class_name = bstr_or_empty!(element.CachedClassName());
    let accelerator_key = bstr_or_empty!(element.CachedAcceleratorKey());

    // Control type -- translate numeric ID to human-readable name.
    let control_type = element
        .CachedControlType()
        .map(|id| control_type_name(id).to_owned())
        .unwrap_or_else(|_| "Unknown".to_owned());

    // Bounding rectangle -- RECT { left, top, right, bottom }.
    let bounding_rect = element
        .CachedBoundingRectangle()
        .map(|r| [r.left as f64, r.top as f64, r.right as f64, r.bottom as f64])
        .unwrap_or([0.0, 0.0, 0.0, 0.0]);

    // Boolean properties.
    let is_offscreen = bool_or_false!(element.CachedIsOffscreen());
    let is_enabled = bool_or_false!(element.CachedIsEnabled());
    let is_control_element = bool_or_false!(element.CachedIsControlElement());
    let has_keyboard_focus = bool_or_false!(element.CachedHasKeyboardFocus());
    let is_keyboard_focusable = bool_or_false!(element.CachedIsKeyboardFocusable());

    // Recurse into children if we have not hit the depth limit.
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

/// Collect children of `parent` from the UIA cache and walk each one.
///
/// Returns an empty `Vec` if `GetCachedChildren` fails (e.g. leaf element,
/// non-accessible window) -- callers must not treat this as an error.
///
/// # Safety
///
/// `parent` must be a valid cached `IUIAutomationElement`.
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
        Ok(n) if n > 0 => n,
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

/// Capture the accessibility tree for one window handle.
///
/// Called from a Rayon worker thread.  Each call:
///
/// 1. Initialises its own MTA COM apartment.
/// 2. Creates a fresh `IUIAutomation`.
/// 3. Calls `ElementFromHandleBuildCache` -- one cross-process RPC that
///    fetches the entire subtree and all requested properties atomically.
/// 4. Walks the cached tree recursively (all local memory, no COM calls).
///
/// Returns `None` when the HWND no longer exists or access is denied --
/// the caller filters these out rather than surfacing errors to Python.
fn capture_window(handle: isize, max_depth: usize) -> Option<TreeElementSnapshot> {
    // Initialise COM for this thread.  The guard is dropped at function exit
    // which calls CoUninitialize.
    let _com_guard = COMGuard::init().ok()?;

    // SAFETY: CoCreateInstance requires COM to be initialised, which
    // COMGuard::init() just did.
    let uia: IUIAutomation = unsafe {
        CoCreateInstance(&CUIAutomation, None, CLSCTX_ALL).ok()?
    };

    // Build the cache request (subtree + all properties).
    // SAFETY: uia is valid and on this thread's apartment.
    let cache_req = unsafe { build_cache_request(&uia).ok()? };

    // Fetch the root element with the full cached subtree.
    // SAFETY: ElementFromHandleBuildCache requires a valid HWND and cache
    // request on the same COM apartment.
    let root: IUIAutomationElement = unsafe {
        uia.ElementFromHandleBuildCache(HWND(handle as *mut core::ffi::c_void), &cache_req)
            .ok()?
    };

    // Walk the cached tree -- all local, no COM calls after this point.
    // SAFETY: root was populated by BuildCache above; all properties and
    // children are in the local cache.
    let snapshot = unsafe { walk_element(&root, 0, max_depth) };

    Some(snapshot)
}

// ---------------------------------------------------------------------------
// Public pyfunction
// ---------------------------------------------------------------------------

/// Capture the Windows UI Automation accessibility tree for one or more
/// windows, returning the result as a list of nested Python dicts.
///
/// # Arguments
///
/// * `window_handles` -- List of window HWNDs as signed integers (matching
///   Python's `int` type for HWND values).
/// * `max_depth` -- Maximum recursion depth (default 50).  Elements deeper
///   than this will have an empty `children` list.
///
/// # Returns
///
/// A `list` of `dict` objects, one per valid window handle.  Windows that
/// no longer exist or that deny access are silently skipped (not an error).
/// Each `dict` has the following keys:
///
/// | Key | Type | Description |
/// |-----|------|-------------|
/// | `name` | `str` | Element display name |
/// | `automation_id` | `str` | Stable programmatic ID |
/// | `control_type` | `str` | Human-readable type (e.g. `"Button"`) |
/// | `localized_control_type` | `str` | Locale-specific type name |
/// | `class_name` | `str` | Win32/WPF class name |
/// | `bounding_rect` | `list[float]` | `[left, top, right, bottom]` px |
/// | `is_offscreen` | `bool` | Not visible on any monitor |
/// | `is_enabled` | `bool` | Not greyed-out / disabled |
/// | `is_control_element` | `bool` | In the logical control tree |
/// | `has_keyboard_focus` | `bool` | Currently focused |
/// | `is_keyboard_focusable` | `bool` | Can receive keyboard focus |
/// | `accelerator_key` | `str` | Keyboard shortcut string |
/// | `depth` | `int` | Depth in tree (root = 0) |
/// | `children` | `list[dict]` | Recursively captured children |
///
/// # Performance
///
/// Uses one `BuildUpdatedCache(TreeScope_Subtree)` call per window -- a
/// single cross-process COM round-trip that fetches the entire subtree and
/// all properties atomically.  Windows are traversed in parallel using Rayon.
/// The GIL is released for the entire COM + traversal phase.
///
/// # Example
///
/// ```python
/// import ctypes, windows_mcp_core
///
/// hwnd = ctypes.windll.user32.GetForegroundWindow()
/// trees = windows_mcp_core.capture_tree([hwnd], max_depth=10)
/// if trees:
///     root = trees[0]
///     print(root["name"], root["control_type"])
///     for child in root["children"]:
///         print(" ", child["name"], child["control_type"])
/// ```
#[pyfunction]
#[pyo3(signature = (window_handles, max_depth=None))]
pub fn capture_tree(
    py: Python<'_>,
    window_handles: Vec<isize>,
    max_depth: Option<usize>,
) -> PyResult<PyObject> {
    // Clamp max_depth to prevent stack overflow on Rayon threads.
    // Real UI trees rarely exceed depth 30.
    let max_depth = max_depth.unwrap_or(50).min(200);

    // Release the GIL for the entire COM traversal phase.  No Python objects
    // are touched inside this closure.
    let snapshots: Vec<TreeElementSnapshot> = py.allow_threads(|| {
        window_handles
            .par_iter()
            .filter_map(|&handle| capture_window(handle, max_depth))
            .collect()
    });

    // Re-acquire the GIL to convert snapshots to Python dicts.
    let result = PyList::empty(py);
    for snapshot in &snapshots {
        result.append(snapshot.to_py_dict(py)?)?;
    }

    Ok(result.into())
}
