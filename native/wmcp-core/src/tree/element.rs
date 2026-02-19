//! Data structures for a single captured accessibility-tree element.
//!
//! [`TreeElementSnapshot`] is an owned, heap-allocated copy of every UIA
//! property read during a `BuildUpdatedCache` pass.  It is fully `Send`
//! and `Serialize` -- no COM references are held.

use serde::Serialize;

/// An owned, COM-free snapshot of one UIA element and its entire subtree.
///
/// All string fields are `String` (UTF-8).  `bounding_rect` stores
/// `[left, top, right, bottom]` as `f64` to match the Python convention.
#[derive(Debug, Clone, Serialize)]
pub struct TreeElementSnapshot {
    pub name: String,
    pub automation_id: String,
    pub control_type: String,
    pub localized_control_type: String,
    pub class_name: String,
    pub bounding_rect: [f64; 4],
    pub is_offscreen: bool,
    pub is_enabled: bool,
    pub is_control_element: bool,
    pub has_keyboard_focus: bool,
    pub is_keyboard_focusable: bool,
    pub accelerator_key: String,
    pub depth: usize,
    pub children: Vec<TreeElementSnapshot>,
}
