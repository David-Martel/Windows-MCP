//! `windows_mcp_core` -- Thin PyO3 wrappers around `wmcp_core`.
//!
//! Each function releases the GIL via `py.allow_threads()` and converts
//! the Rust result to Python objects.  All business logic lives in
//! `wmcp_core`.

use pyo3::exceptions::PyRuntimeError;
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};

use wmcp_core::tree::element::TreeElementSnapshot;

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

    m.add("__version__", env!("CARGO_PKG_VERSION"))?;
    m.add("__doc__", "Native Rust acceleration layer for Windows-MCP.")?;

    Ok(())
}
