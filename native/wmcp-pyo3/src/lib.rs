//! `windows_mcp_core` -- Thin PyO3 wrappers around `wmcp_core`.
//!
//! Each function releases the GIL via `py.allow_threads()` and converts
//! the Rust result to Python objects.  All business logic lives in
//! `wmcp_core`.

use pyo3::exceptions::PyRuntimeError;
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};

use wmcp_core::tree::element::TreeElementSnapshot;

/// Maximum text length accepted by `send_text` (matches core).
const MAX_SEND_TEXT_LEN: usize = 10_000;

/// Maximum window handles accepted by `capture_tree` (matches FFI).
const MAX_HANDLE_COUNT: usize = 256;

// ---------------------------------------------------------------------------
// Error conversion helper
// ---------------------------------------------------------------------------

fn to_py_err(e: wmcp_core::errors::WindowsMcpError) -> PyErr {
    PyRuntimeError::new_err(e.to_string())
}

// ---------------------------------------------------------------------------
// Tree element -> Python dict conversion
// ---------------------------------------------------------------------------

/// Convert a [`TreeElementSnapshot`] tree into a nested Python dict.
///
/// Uses an iterative (stack-based) approach to avoid stack overflow on
/// deeply nested trees, even though Rust caps at `MAX_TREE_DEPTH = 50`.
fn snapshot_to_py_dict(py: Python<'_>, root: &TreeElementSnapshot) -> PyResult<PyObject> {
    // Each stack frame: (snapshot ref, parent PyList to append result to)
    let root_list = PyList::empty(py);
    let mut stack: Vec<(&TreeElementSnapshot, PyObject)> =
        vec![(root, root_list.clone().into())];

    while let Some((snap, parent_list)) = stack.pop() {
        let dict = PyDict::new(py);

        dict.set_item("name", &snap.name)?;
        dict.set_item("automation_id", &snap.automation_id)?;
        dict.set_item("control_type", &snap.control_type)?;
        dict.set_item("localized_control_type", &snap.localized_control_type)?;
        dict.set_item("class_name", &snap.class_name)?;
        dict.set_item("bounding_rect", snap.bounding_rect.to_vec())?;
        dict.set_item("is_offscreen", snap.is_offscreen)?;
        dict.set_item("is_enabled", snap.is_enabled)?;
        dict.set_item("is_control_element", snap.is_control_element)?;
        dict.set_item("has_keyboard_focus", snap.has_keyboard_focus)?;
        dict.set_item("is_keyboard_focusable", snap.is_keyboard_focusable)?;
        dict.set_item("accelerator_key", &snap.accelerator_key)?;
        dict.set_item("depth", snap.depth)?;

        let children_list = PyList::empty(py);
        dict.set_item("children", &children_list)?;

        // Append this dict to the parent's children list
        parent_list.call_method1(py, "append", (dict.as_any(),))?;

        // Push children in reverse so they're processed left-to-right
        for child in snap.children.iter().rev() {
            stack.push((child, children_list.clone().into()));
        }
    }

    // The root_list contains exactly one element (the root dict)
    root_list.get_item(0).map(|item| item.into())
}

/// Convert a [`WindowInfo`] to a Python dict.
fn window_info_to_dict(
    py: Python<'_>,
    info: &wmcp_core::window::WindowInfo,
) -> PyResult<PyObject> {
    let dict = PyDict::new(py);
    dict.set_item("hwnd", info.hwnd)?;
    dict.set_item("title", &info.title)?;
    dict.set_item("class_name", &info.class_name)?;
    dict.set_item("pid", info.pid)?;
    dict.set_item("is_minimized", info.is_minimized)?;
    dict.set_item("is_maximized", info.is_maximized)?;
    dict.set_item("is_visible", info.is_visible)?;

    let rect = PyDict::new(py);
    rect.set_item("left", info.rect.left)?;
    rect.set_item("top", info.rect.top)?;
    rect.set_item("right", info.rect.right)?;
    rect.set_item("bottom", info.rect.bottom)?;
    dict.set_item("rect", rect)?;

    Ok(dict.into())
}

// ---------------------------------------------------------------------------
// system_info
// ---------------------------------------------------------------------------

/// Collect system information and return it as a Python dict.
#[pyfunction]
fn system_info(py: Python<'_>) -> PyResult<PyObject> {
    let snapshot = py
        .allow_threads(wmcp_core::system_info::collect_system_info)
        .map_err(to_py_err)?;

    let dict = PyDict::new(py);

    dict.set_item("os_name", &snapshot.os_name)?;
    dict.set_item("os_version", &snapshot.os_version)?;
    dict.set_item("hostname", &snapshot.hostname)?;
    dict.set_item("cpu_count", snapshot.cpu_count)?;

    let cpu_list = PyList::new(py, snapshot.cpu_usage.iter().map(|&u| u as f64))?;
    dict.set_item("cpu_usage_percent", cpu_list)?;

    dict.set_item("total_memory_bytes", snapshot.total_memory_bytes)?;
    dict.set_item("used_memory_bytes", snapshot.used_memory_bytes)?;

    let disk_list = PyList::empty(py);
    for disk in &snapshot.disks {
        let d = PyDict::new(py);
        d.set_item("name", &disk.name)?;
        d.set_item("mount_point", &disk.mount_point)?;
        d.set_item("total_bytes", disk.total_bytes)?;
        d.set_item("available_bytes", disk.available_bytes)?;
        disk_list.append(d)?;
    }
    dict.set_item("disks", disk_list)?;

    Ok(dict.into())
}

// ---------------------------------------------------------------------------
// capture_tree
// ---------------------------------------------------------------------------

/// Capture the UIA accessibility tree for one or more windows.
#[pyfunction]
#[pyo3(signature = (window_handles, max_depth=None))]
fn capture_tree(
    py: Python<'_>,
    window_handles: Vec<isize>,
    max_depth: Option<usize>,
) -> PyResult<PyObject> {
    if window_handles.len() > MAX_HANDLE_COUNT {
        return Err(PyRuntimeError::new_err(format!(
            "window_handles length {} exceeds maximum {MAX_HANDLE_COUNT}",
            window_handles.len()
        )));
    }

    let max_depth = max_depth.unwrap_or(wmcp_core::tree::MAX_TREE_DEPTH);

    let snapshots = py.allow_threads(|| {
        wmcp_core::tree::capture_tree_raw(&window_handles, max_depth)
    });

    let result = PyList::empty(py);
    for snapshot in &snapshots {
        result.append(snapshot_to_py_dict(py, snapshot)?)?;
    }

    Ok(result.into())
}

// ---------------------------------------------------------------------------
// input functions
// ---------------------------------------------------------------------------

/// Type Unicode text via SendInput.
#[pyfunction]
#[pyo3(signature = (text,))]
fn send_text(py: Python<'_>, text: &str) -> PyResult<u32> {
    if text.len() > MAX_SEND_TEXT_LEN {
        return Err(PyRuntimeError::new_err(format!(
            "text length {} exceeds maximum {MAX_SEND_TEXT_LEN}",
            text.len()
        )));
    }
    let text_owned = text.to_owned();
    Ok(py.allow_threads(move || wmcp_core::input::send_text_raw(&text_owned)))
}

/// Press or release a virtual key code.
#[pyfunction]
#[pyo3(signature = (vk_code, key_up=false))]
fn send_key(py: Python<'_>, vk_code: u16, key_up: bool) -> PyResult<u32> {
    Ok(py.allow_threads(move || wmcp_core::input::send_key_raw(vk_code, key_up)))
}

/// Click the mouse at absolute screen coordinates.
#[pyfunction]
#[pyo3(signature = (x, y, button="left"))]
fn send_click(py: Python<'_>, x: i32, y: i32, button: &str) -> PyResult<u32> {
    let button_owned = button.to_lowercase();
    Ok(py.allow_threads(move || wmcp_core::input::send_click_raw(x, y, &button_owned)))
}

/// Move the mouse cursor to absolute screen coordinates.
#[pyfunction]
#[pyo3(signature = (x, y))]
fn send_mouse_move(py: Python<'_>, x: i32, y: i32) -> PyResult<u32> {
    Ok(py.allow_threads(move || wmcp_core::input::send_mouse_move_raw(x, y)))
}

/// Send a key combination (e.g. Ctrl+C).
#[pyfunction]
#[pyo3(signature = (vk_codes,))]
fn send_hotkey(py: Python<'_>, vk_codes: Vec<u16>) -> PyResult<u32> {
    Ok(py.allow_threads(move || wmcp_core::input::send_hotkey_raw(&vk_codes)))
}

/// Scroll the mouse wheel at screen coordinates.
#[pyfunction]
#[pyo3(signature = (x, y, delta, horizontal=false))]
fn send_scroll(py: Python<'_>, x: i32, y: i32, delta: i32, horizontal: bool) -> PyResult<u32> {
    Ok(py.allow_threads(move || wmcp_core::input::send_scroll_raw(x, y, delta, horizontal)))
}

/// Drag the mouse from current position to destination coordinates.
#[pyfunction]
#[pyo3(signature = (to_x, to_y, steps=10))]
fn send_drag(py: Python<'_>, to_x: i32, to_y: i32, steps: u32) -> PyResult<u32> {
    Ok(py.allow_threads(move || wmcp_core::input::send_drag_raw(to_x, to_y, steps)))
}

// ---------------------------------------------------------------------------
// window functions
// ---------------------------------------------------------------------------

/// Enumerate all visible top-level windows (Alt+Tab windows with titles).
#[pyfunction]
fn enumerate_windows(py: Python<'_>) -> PyResult<PyObject> {
    let handles = py
        .allow_threads(wmcp_core::window::enumerate_visible_windows)
        .map_err(to_py_err)?;

    Ok(PyList::new(py, &handles)?.into())
}

/// Get detailed information about a window by handle.
#[pyfunction]
#[pyo3(signature = (hwnd,))]
fn get_window_info(py: Python<'_>, hwnd: isize) -> PyResult<PyObject> {
    let info = py
        .allow_threads(move || wmcp_core::window::get_window_info(hwnd))
        .map_err(to_py_err)?;

    window_info_to_dict(py, &info)
}

/// Get the foreground (active) window handle.
#[pyfunction]
fn get_foreground_window(py: Python<'_>) -> PyResult<isize> {
    Ok(py.allow_threads(wmcp_core::window::get_foreground_hwnd))
}

/// List all visible windows with their information.
#[pyfunction]
fn list_windows(py: Python<'_>) -> PyResult<PyObject> {
    let windows = py
        .allow_threads(wmcp_core::window::list_windows)
        .map_err(to_py_err)?;

    let result = PyList::empty(py);
    for info in &windows {
        result.append(window_info_to_dict(py, info)?)?;
    }

    Ok(result.into())
}

// ---------------------------------------------------------------------------
// screenshot functions
// ---------------------------------------------------------------------------

/// Capture a screenshot as raw BGRA pixel bytes.
///
/// Returns a dict with keys: `width` (int), `height` (int), `data` (bytes).
#[pyfunction]
#[pyo3(signature = (monitor_index=0))]
fn capture_screenshot_raw(py: Python<'_>, monitor_index: u32) -> PyResult<PyObject> {
    let frame = py
        .allow_threads(move || wmcp_core::screenshot::capture_raw(monitor_index))
        .map_err(to_py_err)?;

    let dict = PyDict::new(py);
    dict.set_item("width", frame.width)?;
    dict.set_item("height", frame.height)?;
    dict.set_item("data", pyo3::types::PyBytes::new(py, &frame.data))?;
    Ok(dict.into())
}

/// Capture a screenshot and encode it as PNG bytes.
///
/// Returns a `bytes` object containing the PNG file data.
#[pyfunction]
#[pyo3(signature = (monitor_index=0))]
fn capture_screenshot_png(py: Python<'_>, monitor_index: u32) -> PyResult<PyObject> {
    let png_bytes = py
        .allow_threads(move || wmcp_core::screenshot::capture_png(monitor_index))
        .map_err(to_py_err)?;

    Ok(pyo3::types::PyBytes::new(py, &png_bytes).into())
}

// ---------------------------------------------------------------------------
// UIA query functions
// ---------------------------------------------------------------------------

/// Convert an [`ElementInfo`] to a Python dict.
fn element_info_to_dict(
    py: Python<'_>,
    info: &wmcp_core::query::ElementInfo,
) -> PyResult<PyObject> {
    let dict = PyDict::new(py);
    dict.set_item("name", &info.name)?;
    dict.set_item("automation_id", &info.automation_id)?;
    dict.set_item("control_type", &info.control_type)?;
    dict.set_item("localized_control_type", &info.localized_control_type)?;
    dict.set_item("class_name", &info.class_name)?;
    dict.set_item("bounding_rect", info.bounding_rect.to_vec())?;
    dict.set_item("is_enabled", info.is_enabled)?;
    dict.set_item("is_offscreen", info.is_offscreen)?;
    dict.set_item("has_keyboard_focus", info.has_keyboard_focus)?;
    dict.set_item("supported_patterns", &info.supported_patterns)?;
    Ok(dict.into())
}

/// Convert a [`PatternResult`] to a Python dict.
fn pattern_result_to_dict(
    py: Python<'_>,
    r: &wmcp_core::pattern::PatternResult,
) -> PyResult<PyObject> {
    let dict = PyDict::new(py);
    dict.set_item("element_name", &r.element_name)?;
    dict.set_item("element_type", &r.element_type)?;
    dict.set_item("action", &r.action)?;
    dict.set_item("success", r.success)?;
    dict.set_item("detail", &r.detail)?;
    Ok(dict.into())
}

/// Query the UIA element at screen coordinates.
#[pyfunction]
#[pyo3(signature = (x, y))]
fn element_from_point(py: Python<'_>, x: i32, y: i32) -> PyResult<PyObject> {
    let info = py
        .allow_threads(move || wmcp_core::query::element_from_point(x, y))
        .map_err(to_py_err)?;
    element_info_to_dict(py, &info)
}

/// Search for UIA elements matching criteria.
#[pyfunction]
#[pyo3(signature = (name=None, control_type=None, automation_id=None, window_handle=None, limit=20))]
fn find_elements(
    py: Python<'_>,
    name: Option<String>,
    control_type: Option<String>,
    automation_id: Option<String>,
    window_handle: Option<isize>,
    limit: usize,
) -> PyResult<PyObject> {
    let criteria = wmcp_core::query::FindCriteria {
        name,
        control_type,
        automation_id,
        window_handle,
        limit,
    };

    let results = py
        .allow_threads(move || wmcp_core::query::find_elements(&criteria))
        .map_err(to_py_err)?;

    let list = PyList::empty(py);
    for info in &results {
        list.append(element_info_to_dict(py, info)?)?;
    }
    Ok(list.into())
}

/// Query primary and virtual screen dimensions.
#[pyfunction]
fn get_screen_metrics(py: Python<'_>) -> PyResult<PyObject> {
    let metrics = py
        .allow_threads(wmcp_core::query::get_screen_metrics)
        .map_err(to_py_err)?;

    let dict = PyDict::new(py);
    dict.set_item("primary_width", metrics.primary_width)?;
    dict.set_item("primary_height", metrics.primary_height)?;
    dict.set_item("virtual_width", metrics.virtual_width)?;
    dict.set_item("virtual_height", metrics.virtual_height)?;
    Ok(dict.into())
}

// ---------------------------------------------------------------------------
// UIA pattern functions
// ---------------------------------------------------------------------------

/// Invoke the InvokePattern on the element at (x, y).
#[pyfunction]
#[pyo3(signature = (x, y))]
fn invoke_at(py: Python<'_>, x: i32, y: i32) -> PyResult<PyObject> {
    let result = py
        .allow_threads(move || wmcp_core::pattern::invoke_at(x, y))
        .map_err(to_py_err)?;
    pattern_result_to_dict(py, &result)
}

/// Toggle the TogglePattern on the element at (x, y).
#[pyfunction]
#[pyo3(signature = (x, y))]
fn toggle_at(py: Python<'_>, x: i32, y: i32) -> PyResult<PyObject> {
    let result = py
        .allow_threads(move || wmcp_core::pattern::toggle_at(x, y))
        .map_err(to_py_err)?;
    pattern_result_to_dict(py, &result)
}

/// Set a value via ValuePattern on the element at (x, y).
#[pyfunction]
#[pyo3(signature = (x, y, value))]
fn set_value_at(py: Python<'_>, x: i32, y: i32, value: &str) -> PyResult<PyObject> {
    let value_owned = value.to_owned();
    let result = py
        .allow_threads(move || wmcp_core::pattern::set_value_at(x, y, &value_owned))
        .map_err(to_py_err)?;
    pattern_result_to_dict(py, &result)
}

/// Expand via ExpandCollapsePattern on the element at (x, y).
#[pyfunction]
#[pyo3(signature = (x, y))]
fn expand_at(py: Python<'_>, x: i32, y: i32) -> PyResult<PyObject> {
    let result = py
        .allow_threads(move || wmcp_core::pattern::expand_at(x, y))
        .map_err(to_py_err)?;
    pattern_result_to_dict(py, &result)
}

/// Collapse via ExpandCollapsePattern on the element at (x, y).
#[pyfunction]
#[pyo3(signature = (x, y))]
fn collapse_at(py: Python<'_>, x: i32, y: i32) -> PyResult<PyObject> {
    let result = py
        .allow_threads(move || wmcp_core::pattern::collapse_at(x, y))
        .map_err(to_py_err)?;
    pattern_result_to_dict(py, &result)
}

/// Select via SelectionItemPattern on the element at (x, y).
#[pyfunction]
#[pyo3(signature = (x, y))]
fn select_at(py: Python<'_>, x: i32, y: i32) -> PyResult<PyObject> {
    let result = py
        .allow_threads(move || wmcp_core::pattern::select_at(x, y))
        .map_err(to_py_err)?;
    pattern_result_to_dict(py, &result)
}

// ---------------------------------------------------------------------------
// Module registration
// ---------------------------------------------------------------------------

/// Register the `windows_mcp_core` Python module.
#[pymodule]
fn windows_mcp_core(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(system_info, m)?)?;
    m.add_function(wrap_pyfunction!(capture_tree, m)?)?;
    m.add_function(wrap_pyfunction!(send_text, m)?)?;
    m.add_function(wrap_pyfunction!(send_key, m)?)?;
    m.add_function(wrap_pyfunction!(send_click, m)?)?;
    m.add_function(wrap_pyfunction!(send_mouse_move, m)?)?;
    m.add_function(wrap_pyfunction!(send_hotkey, m)?)?;
    m.add_function(wrap_pyfunction!(send_scroll, m)?)?;
    m.add_function(wrap_pyfunction!(send_drag, m)?)?;
    m.add_function(wrap_pyfunction!(enumerate_windows, m)?)?;
    m.add_function(wrap_pyfunction!(get_window_info, m)?)?;
    m.add_function(wrap_pyfunction!(get_foreground_window, m)?)?;
    m.add_function(wrap_pyfunction!(list_windows, m)?)?;
    m.add_function(wrap_pyfunction!(capture_screenshot_raw, m)?)?;
    m.add_function(wrap_pyfunction!(capture_screenshot_png, m)?)?;
    // UIA query functions
    m.add_function(wrap_pyfunction!(element_from_point, m)?)?;
    m.add_function(wrap_pyfunction!(find_elements, m)?)?;
    m.add_function(wrap_pyfunction!(get_screen_metrics, m)?)?;
    // UIA pattern functions
    m.add_function(wrap_pyfunction!(invoke_at, m)?)?;
    m.add_function(wrap_pyfunction!(toggle_at, m)?)?;
    m.add_function(wrap_pyfunction!(set_value_at, m)?)?;
    m.add_function(wrap_pyfunction!(expand_at, m)?)?;
    m.add_function(wrap_pyfunction!(collapse_at, m)?)?;
    m.add_function(wrap_pyfunction!(select_at, m)?)?;

    m.add("__version__", env!("CARGO_PKG_VERSION"))?;
    m.add("__doc__", "Native Rust acceleration layer for Windows-MCP.")?;

    Ok(())
}
