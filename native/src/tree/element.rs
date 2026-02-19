//! Data structures for a single captured accessibility-tree element.
//!
//! [`TreeElementSnapshot`] is an owned, heap-allocated copy of every UIA
//! property we read during a single `BuildUpdatedCache` pass.  Because all
//! COM calls happen before we re-acquire the GIL, the snapshot is fully
//! `Send` and can be converted to a Python `dict` without holding any COM
//! references.
//!
//! # Memory layout
//!
//! A tree snapshot forms a recursive structure:
//!
//! ```text
//! TreeElementSnapshot {
//!     name: "Save",
//!     children: [
//!         TreeElementSnapshot { name: "Save icon", children: [], … },
//!     ],
//!     …
//! }
//! ```
//!
//! The `to_py_dict` method walks this recursively, building nested Python
//! `dict` / `list` objects.

use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};

/// An owned, COM-free snapshot of one UIA element and its entire subtree.
///
/// All string fields are `String` (UTF-8) rather than `windows_core::BSTR`
/// so the struct is `Send` and can cross thread boundaries safely.
///
/// `bounding_rect` stores `[left, top, right, bottom]` as `f64` to match
/// the Python convention used by the existing `tree/views.py`.
#[derive(Debug, Clone)]
pub struct TreeElementSnapshot {
    /// `UIA_NamePropertyId` -- display name of the element.
    pub name: String,

    /// `UIA_AutomationIdPropertyId` -- stable programmatic identifier.
    pub automation_id: String,

    /// Human-readable control type name resolved from `UIA_ControlTypePropertyId`.
    /// Example: `"Button"`, `"Edit"`, `"Tree"`.
    pub control_type: String,

    /// `UIA_LocalizedControlTypePropertyId` -- locale-specific control name.
    pub localized_control_type: String,

    /// `UIA_ClassNamePropertyId` -- Win32/WPF class name.
    pub class_name: String,

    /// `UIA_BoundingRectanglePropertyId` -- `[left, top, right, bottom]` in
    /// screen coordinates (pixels).  All zeros when the element is off-screen
    /// or the property is unavailable.
    pub bounding_rect: [f64; 4],

    /// `UIA_IsOffscreenPropertyId` -- element is not visible on any monitor.
    pub is_offscreen: bool,

    /// `UIA_IsEnabledPropertyId` -- element is not greyed-out / disabled.
    pub is_enabled: bool,

    /// `UIA_IsControlElementPropertyId` -- element participates in logical
    /// tree (vs. raw / content tree).
    pub is_control_element: bool,

    /// `UIA_HasKeyboardFocusPropertyId` -- element currently has focus.
    pub has_keyboard_focus: bool,

    /// `UIA_IsKeyboardFocusablePropertyId` -- element can receive keyboard
    /// focus.
    pub is_keyboard_focusable: bool,

    /// `UIA_AcceleratorKeyPropertyId` -- keyboard shortcut string.
    pub accelerator_key: String,

    /// Depth in the accessibility tree (root = 0).  Useful for indented
    /// display and for enforcing `max_depth` on the Python side.
    pub depth: usize,

    /// Recursively captured child elements (populated when depth <
    /// `max_depth`).
    pub children: Vec<TreeElementSnapshot>,
}

impl TreeElementSnapshot {
    /// Convert this snapshot to a Python `dict` recursively.
    ///
    /// Children are converted to a Python `list` of `dict` objects stored
    /// under the `"children"` key.
    ///
    /// # Errors
    ///
    /// Returns `PyErr` only on OOM or internal PyO3 failures -- both are
    /// unrecoverable in practice.
    ///
    /// # Example
    ///
    /// ```python
    /// import windows_mcp_core
    ///
    /// results = windows_mcp_core.capture_tree([hwnd], max_depth=10)
    /// root = results[0]            # dict for the window root element
    /// print(root["name"])          # e.g. "Notepad"
    /// print(root["control_type"])  # e.g. "Window"
    /// children = root["children"]  # list of dicts
    /// ```
    pub fn to_py_dict(&self, py: Python<'_>) -> PyResult<PyObject> {
        let dict = PyDict::new(py);

        dict.set_item("name", &self.name)?;
        dict.set_item("automation_id", &self.automation_id)?;
        dict.set_item("control_type", &self.control_type)?;
        dict.set_item("localized_control_type", &self.localized_control_type)?;
        dict.set_item("class_name", &self.class_name)?;
        dict.set_item("bounding_rect", self.bounding_rect.to_vec())?;
        dict.set_item("is_offscreen", self.is_offscreen)?;
        dict.set_item("is_enabled", self.is_enabled)?;
        dict.set_item("is_control_element", self.is_control_element)?;
        dict.set_item("has_keyboard_focus", self.has_keyboard_focus)?;
        dict.set_item("is_keyboard_focusable", self.is_keyboard_focusable)?;
        dict.set_item("accelerator_key", &self.accelerator_key)?;
        dict.set_item("depth", self.depth)?;

        // Recursively convert children.
        let children_list = PyList::empty(py);
        for child in &self.children {
            children_list.append(child.to_py_dict(py)?)?;
        }
        dict.set_item("children", children_list)?;

        Ok(dict.into())
    }
}
