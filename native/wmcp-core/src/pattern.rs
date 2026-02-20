//! UIA pattern invocation: Invoke, Toggle, SetValue, Expand, Collapse, Select.
//!
//! Each function locates the element at screen coordinates via `ElementFromPoint`,
//! then invokes the requested UIA pattern.  All functions are pure Rust with no
//! PyO3 dependency.
//!
//! # COM apartment model
//!
//! Each function initialises its own MTA COM apartment via [`COMGuard`].

use serde::Serialize;
use windows::core::Interface;
use windows::Win32::Foundation::POINT;
use windows::Win32::System::Com::{CoCreateInstance, CLSCTX_INPROC_SERVER};
use windows::Win32::UI::Accessibility::{
    CUIAutomation, IUIAutomation, IUIAutomationElement, IUIAutomationExpandCollapsePattern,
    IUIAutomationInvokePattern, IUIAutomationSelectionItemPattern, IUIAutomationTogglePattern,
    IUIAutomationValuePattern, UIA_ExpandCollapsePatternId, UIA_InvokePatternId,
    UIA_SelectionItemPatternId, UIA_TogglePatternId, UIA_ValuePatternId,
};

use crate::com::COMGuard;
use crate::errors::WindowsMcpError;
use crate::tree::control_type_name;

// ---------------------------------------------------------------------------
// Data structures
// ---------------------------------------------------------------------------

/// Result of a UIA pattern invocation.
#[derive(Debug, Clone, Serialize)]
pub struct PatternResult {
    pub element_name: String,
    pub element_type: String,
    pub action: String,
    pub success: bool,
    pub detail: String,
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/// Locate the UIA element at screen coordinates.
///
/// Returns `(IUIAutomation, IUIAutomationElement)` so the caller can use the
/// same UIA instance for pattern queries.
unsafe fn element_at(
    x: i32,
    y: i32,
) -> Result<(IUIAutomation, IUIAutomationElement), WindowsMcpError> {
    let uia: IUIAutomation = CoCreateInstance(&CUIAutomation, None, CLSCTX_INPROC_SERVER)?;

    let point = POINT { x, y };
    let element = uia
        .ElementFromPoint(point)
        .map_err(|e| WindowsMcpError::TreeError(format!("ElementFromPoint({x},{y}): {e}")))?;

    Ok((uia, element))
}

/// Read element name for diagnostics.
unsafe fn elem_name(element: &IUIAutomationElement) -> String {
    element
        .CurrentName()
        .map(|b| b.to_string())
        .unwrap_or_default()
}

/// Read element localized control type for diagnostics.
unsafe fn elem_type(element: &IUIAutomationElement) -> String {
    element
        .CurrentControlType()
        .map(|id| control_type_name(id).to_owned())
        .unwrap_or_else(|_| "Unknown".to_owned())
}

/// Build a [`PatternResult`] with `success = false`.
fn pattern_not_supported(name: &str, etype: &str, action: &str, pattern_name: &str) -> PatternResult {
    PatternResult {
        element_name: name.to_owned(),
        element_type: etype.to_owned(),
        action: action.to_owned(),
        success: false,
        detail: format!("Element does not support {pattern_name}"),
    }
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/// Invoke the `InvokePattern` on the element at `(x, y)`.
pub fn invoke_at(x: i32, y: i32) -> Result<PatternResult, WindowsMcpError> {
    let _com = COMGuard::init()?;

    let (_uia, element) = unsafe { element_at(x, y)? };
    let name = unsafe { elem_name(&element) };
    let etype = unsafe { elem_type(&element) };

    let pattern: Option<IUIAutomationInvokePattern> = unsafe {
        element
            .GetCurrentPattern(UIA_InvokePatternId)
            .ok()
            .and_then(|p| p.cast::<IUIAutomationInvokePattern>().ok())
    };

    match pattern {
        Some(p) => {
            unsafe { p.Invoke() }
                .map_err(|e| WindowsMcpError::TreeError(format!("Invoke failed: {e}")))?;
            Ok(PatternResult {
                element_name: name,
                element_type: etype,
                action: "invoke".into(),
                success: true,
                detail: format!("Invoked at ({x},{y})"),
            })
        }
        None => Ok(pattern_not_supported(&name, &etype, "invoke", "InvokePattern")),
    }
}

/// Toggle the `TogglePattern` on the element at `(x, y)`.
///
/// Returns the new toggle state in `detail` (e.g. "State: on").
pub fn toggle_at(x: i32, y: i32) -> Result<PatternResult, WindowsMcpError> {
    let _com = COMGuard::init()?;

    let (_uia, element) = unsafe { element_at(x, y)? };
    let name = unsafe { elem_name(&element) };
    let etype = unsafe { elem_type(&element) };

    let pattern: Option<IUIAutomationTogglePattern> = unsafe {
        element
            .GetCurrentPattern(UIA_TogglePatternId)
            .ok()
            .and_then(|p| p.cast::<IUIAutomationTogglePattern>().ok())
    };

    match pattern {
        Some(p) => {
            unsafe { p.Toggle() }
                .map_err(|e| WindowsMcpError::TreeError(format!("Toggle failed: {e}")))?;

            let state = unsafe { p.CurrentToggleState() }.unwrap_or_default();
            let state_name = match state.0 {
                0 => "off",
                1 => "on",
                2 => "indeterminate",
                _ => "unknown",
            };

            Ok(PatternResult {
                element_name: name,
                element_type: etype,
                action: "toggle".into(),
                success: true,
                detail: format!("State: {state_name}"),
            })
        }
        None => Ok(pattern_not_supported(&name, &etype, "toggle", "TogglePattern")),
    }
}

/// Set a value via `ValuePattern` on the element at `(x, y)`.
pub fn set_value_at(x: i32, y: i32, value: &str) -> Result<PatternResult, WindowsMcpError> {
    let _com = COMGuard::init()?;

    let (_uia, element) = unsafe { element_at(x, y)? };
    let name = unsafe { elem_name(&element) };
    let etype = unsafe { elem_type(&element) };

    let pattern: Option<IUIAutomationValuePattern> = unsafe {
        element
            .GetCurrentPattern(UIA_ValuePatternId)
            .ok()
            .and_then(|p| p.cast::<IUIAutomationValuePattern>().ok())
    };

    match pattern {
        Some(p) => {
            let bstr = windows::core::BSTR::from(value);
            unsafe { p.SetValue(&bstr) }
                .map_err(|e| WindowsMcpError::TreeError(format!("SetValue failed: {e}")))?;

            let preview = if value.len() > 50 {
                format!("{}...", &value[..50])
            } else {
                value.to_owned()
            };

            Ok(PatternResult {
                element_name: name,
                element_type: etype,
                action: "set_value".into(),
                success: true,
                detail: format!("Value set to '{preview}'"),
            })
        }
        None => Ok(pattern_not_supported(&name, &etype, "set_value", "ValuePattern")),
    }
}

/// Expand via `ExpandCollapsePattern` on the element at `(x, y)`.
pub fn expand_at(x: i32, y: i32) -> Result<PatternResult, WindowsMcpError> {
    let _com = COMGuard::init()?;

    let (_uia, element) = unsafe { element_at(x, y)? };
    let name = unsafe { elem_name(&element) };
    let etype = unsafe { elem_type(&element) };

    let pattern: Option<IUIAutomationExpandCollapsePattern> = unsafe {
        element
            .GetCurrentPattern(UIA_ExpandCollapsePatternId)
            .ok()
            .and_then(|p| p.cast::<IUIAutomationExpandCollapsePattern>().ok())
    };

    match pattern {
        Some(p) => {
            unsafe { p.Expand() }
                .map_err(|e| WindowsMcpError::TreeError(format!("Expand failed: {e}")))?;
            Ok(PatternResult {
                element_name: name,
                element_type: etype,
                action: "expand".into(),
                success: true,
                detail: format!("Expanded at ({x},{y})"),
            })
        }
        None => Ok(pattern_not_supported(
            &name,
            &etype,
            "expand",
            "ExpandCollapsePattern",
        )),
    }
}

/// Collapse via `ExpandCollapsePattern` on the element at `(x, y)`.
pub fn collapse_at(x: i32, y: i32) -> Result<PatternResult, WindowsMcpError> {
    let _com = COMGuard::init()?;

    let (_uia, element) = unsafe { element_at(x, y)? };
    let name = unsafe { elem_name(&element) };
    let etype = unsafe { elem_type(&element) };

    let pattern: Option<IUIAutomationExpandCollapsePattern> = unsafe {
        element
            .GetCurrentPattern(UIA_ExpandCollapsePatternId)
            .ok()
            .and_then(|p| p.cast::<IUIAutomationExpandCollapsePattern>().ok())
    };

    match pattern {
        Some(p) => {
            unsafe { p.Collapse() }
                .map_err(|e| WindowsMcpError::TreeError(format!("Collapse failed: {e}")))?;
            Ok(PatternResult {
                element_name: name,
                element_type: etype,
                action: "collapse".into(),
                success: true,
                detail: format!("Collapsed at ({x},{y})"),
            })
        }
        None => Ok(pattern_not_supported(
            &name,
            &etype,
            "collapse",
            "ExpandCollapsePattern",
        )),
    }
}

/// Select via `SelectionItemPattern` on the element at `(x, y)`.
pub fn select_at(x: i32, y: i32) -> Result<PatternResult, WindowsMcpError> {
    let _com = COMGuard::init()?;

    let (_uia, element) = unsafe { element_at(x, y)? };
    let name = unsafe { elem_name(&element) };
    let etype = unsafe { elem_type(&element) };

    let pattern: Option<IUIAutomationSelectionItemPattern> = unsafe {
        element
            .GetCurrentPattern(UIA_SelectionItemPatternId)
            .ok()
            .and_then(|p| p.cast::<IUIAutomationSelectionItemPattern>().ok())
    };

    match pattern {
        Some(p) => {
            unsafe { p.Select() }
                .map_err(|e| WindowsMcpError::TreeError(format!("Select failed: {e}")))?;
            Ok(PatternResult {
                element_name: name,
                element_type: etype,
                action: "select".into(),
                success: true,
                detail: format!("Selected at ({x},{y})"),
            })
        }
        None => Ok(pattern_not_supported(
            &name,
            &etype,
            "select",
            "SelectionItemPattern",
        )),
    }
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_pattern_result_serialization() {
        let r = PatternResult {
            element_name: "OK Button".into(),
            element_type: "Button".into(),
            action: "invoke".into(),
            success: true,
            detail: "Invoked at (100,200)".into(),
        };
        let json = serde_json::to_string(&r).unwrap();
        assert!(json.contains("\"success\":true"));
        assert!(json.contains("OK Button"));
    }

    #[test]
    fn test_pattern_result_failure() {
        let r = pattern_not_supported("test", "Button", "toggle", "TogglePattern");
        assert!(!r.success);
        assert!(r.detail.contains("TogglePattern"));
    }

    #[test]
    fn test_pattern_result_detail_formatting() {
        let r = PatternResult {
            element_name: "Check".into(),
            element_type: "CheckBox".into(),
            action: "toggle".into(),
            success: true,
            detail: "State: on".into(),
        };
        assert_eq!(r.detail, "State: on");
    }

    #[test]
    fn test_set_value_preview_truncation() {
        let long_value = "a".repeat(100);
        let preview = if long_value.len() > 50 {
            format!("{}...", &long_value[..50])
        } else {
            long_value.clone()
        };
        assert_eq!(preview.len(), 53); // 50 chars + "..."
        assert!(preview.ends_with("..."));
    }
}
