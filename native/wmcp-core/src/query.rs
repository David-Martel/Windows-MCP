//! UIA element queries: `ElementFromPoint`, `FindFirst`/`FindAll`, screen metrics.
//!
//! All functions are pure Rust with no PyO3 dependency.  PyO3 wrappers
//! in `wmcp-pyo3` call these via `py.allow_threads()`.
//!
//! # COM apartment model
//!
//! Each function initialises its own MTA COM apartment via [`COMGuard`].
//! COM interfaces are never shared across function boundaries.

use serde::Serialize;
use windows::core::Interface;
use windows::Win32::Foundation::{HWND, POINT};
use windows::Win32::System::Com::{CoCreateInstance, CLSCTX_INPROC_SERVER};
use windows::Win32::UI::Accessibility::{
    CUIAutomation, IUIAutomation, IUIAutomationCondition, IUIAutomationElement,
    TreeScope_Descendants, UIA_AutomationIdPropertyId, UIA_ControlTypePropertyId,
    UIA_ExpandCollapsePatternId, UIA_InvokePatternId, UIA_SelectionItemPatternId,
    UIA_TogglePatternId, UIA_ValuePatternId,
};
use windows::Win32::UI::WindowsAndMessaging::{
    GetSystemMetrics, SM_CXSCREEN, SM_CXVIRTUALSCREEN, SM_CYSCREEN, SM_CYVIRTUALSCREEN,
};

use crate::com::COMGuard;
use crate::errors::WindowsMcpError;
use crate::tree::control_type_name;

/// Maximum number of results from `find_elements`.
const MAX_FIND_LIMIT: usize = 100;

/// UIA pattern IDs to probe for `supported_patterns`.
///
/// Stores the raw i32 pattern IDs (used with `GetCurrentPattern` which
/// takes `UIA_PATTERN_ID` -- a newtype around i32).
const PATTERN_PROBES: &[(i32, &str)] = &[
    (UIA_InvokePatternId.0, "InvokePattern"),
    (UIA_TogglePatternId.0, "TogglePattern"),
    (UIA_ValuePatternId.0, "ValuePattern"),
    (UIA_ExpandCollapsePatternId.0, "ExpandCollapsePattern"),
    (UIA_SelectionItemPatternId.0, "SelectionItemPattern"),
];

// ---------------------------------------------------------------------------
// Data structures
// ---------------------------------------------------------------------------

/// An owned, COM-free snapshot of a single UIA element's properties.
///
/// All strings are UTF-8.  `bounding_rect` stores `[left, top, right, bottom]`
/// as `f64` to match the Python convention.
#[derive(Debug, Clone, Serialize)]
pub struct ElementInfo {
    pub name: String,
    pub automation_id: String,
    pub control_type: String,
    pub localized_control_type: String,
    pub class_name: String,
    pub bounding_rect: [f64; 4],
    pub is_enabled: bool,
    pub is_offscreen: bool,
    pub has_keyboard_focus: bool,
    pub supported_patterns: Vec<String>,
}

/// Criteria for [`find_elements`].
#[derive(Debug, Clone, Default)]
pub struct FindCriteria {
    /// Substring match on element name (case-insensitive).
    pub name: Option<String>,
    /// Exact match on control type name (e.g. "Button").
    pub control_type: Option<String>,
    /// Exact match on AutomationId.
    pub automation_id: Option<String>,
    /// Scope search to a specific window handle.
    pub window_handle: Option<isize>,
    /// Maximum results (clamped to [`MAX_FIND_LIMIT`]).
    pub limit: usize,
}

/// Primary and virtual screen dimensions.
#[derive(Debug, Clone, Serialize)]
pub struct ScreenMetrics {
    pub primary_width: i32,
    pub primary_height: i32,
    pub virtual_width: i32,
    pub virtual_height: i32,
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/// Read common properties from a live UIA element into an owned [`ElementInfo`].
unsafe fn read_element_info(element: &IUIAutomationElement) -> ElementInfo {
    let name = element
        .CurrentName()
        .map(|b| b.to_string())
        .unwrap_or_default();
    let automation_id = element
        .CurrentAutomationId()
        .map(|b| b.to_string())
        .unwrap_or_default();
    let control_type = element
        .CurrentControlType()
        .map(|id| control_type_name(id).to_owned())
        .unwrap_or_else(|_| "Unknown".to_owned());
    let localized_control_type = element
        .CurrentLocalizedControlType()
        .map(|b| b.to_string())
        .unwrap_or_default();
    let class_name = element
        .CurrentClassName()
        .map(|b| b.to_string())
        .unwrap_or_default();

    let bounding_rect = element
        .CurrentBoundingRectangle()
        .map(|r| [r.left as f64, r.top as f64, r.right as f64, r.bottom as f64])
        .unwrap_or([0.0, 0.0, 0.0, 0.0]);

    let is_enabled = element
        .CurrentIsEnabled()
        .map(|b| b.as_bool())
        .unwrap_or(false);
    let is_offscreen = element
        .CurrentIsOffscreen()
        .map(|b| b.as_bool())
        .unwrap_or(false);
    let has_keyboard_focus = element
        .CurrentHasKeyboardFocus()
        .map(|b| b.as_bool())
        .unwrap_or(false);

    // Probe supported patterns -- GetCurrentPattern returns Err if unsupported
    let mut supported_patterns = Vec::new();
    for &(pattern_id, pattern_name) in PATTERN_PROBES {
        use windows::Win32::UI::Accessibility::UIA_PATTERN_ID;
        if element
            .GetCurrentPattern(UIA_PATTERN_ID(pattern_id))
            .is_ok()
        {
            supported_patterns.push(pattern_name.to_owned());
        }
    }

    ElementInfo {
        name,
        automation_id,
        control_type,
        localized_control_type,
        class_name,
        bounding_rect,
        is_enabled,
        is_offscreen,
        has_keyboard_focus,
        supported_patterns,
    }
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/// Query the UIA element at the given screen coordinates.
///
/// Returns an [`ElementInfo`] with all commonly needed properties, or an
/// error if no element is found or COM fails.
pub fn element_from_point(x: i32, y: i32) -> Result<ElementInfo, WindowsMcpError> {
    let _com = COMGuard::init()?;

    let uia: IUIAutomation = unsafe {
        CoCreateInstance(&CUIAutomation, None, CLSCTX_INPROC_SERVER)?
    };

    let point = POINT { x, y };
    let element: IUIAutomationElement = unsafe {
        uia.ElementFromPoint(point)
            .map_err(|e| WindowsMcpError::TreeError(format!("ElementFromPoint({x},{y}): {e}")))?
    };

    let info = unsafe { read_element_info(&element) };
    Ok(info)
}

/// Search for UIA elements matching the given criteria.
///
/// If `criteria.window_handle` is set, the search is scoped to that window's
/// subtree.  Otherwise, the desktop root element is used.
///
/// Returns up to `criteria.limit` matches (clamped to [`MAX_FIND_LIMIT`]).
pub fn find_elements(criteria: &FindCriteria) -> Result<Vec<ElementInfo>, WindowsMcpError> {
    let _com = COMGuard::init()?;

    let uia: IUIAutomation = unsafe {
        CoCreateInstance(&CUIAutomation, None, CLSCTX_INPROC_SERVER)?
    };

    // Determine root element (window or desktop)
    let root: IUIAutomationElement = unsafe {
        if let Some(hwnd) = criteria.window_handle {
            uia.ElementFromHandle(HWND(hwnd as *mut core::ffi::c_void))
                .map_err(|e| WindowsMcpError::TreeError(format!("ElementFromHandle: {e}")))?
        } else {
            uia.GetRootElement()
                .map_err(|e| WindowsMcpError::TreeError(format!("GetRootElement: {e}")))?
        }
    };

    // Build condition
    let condition: IUIAutomationCondition = unsafe {
        build_find_condition(&uia, criteria)?
    };

    // FindAll with TreeScope_Descendants
    let elements = unsafe {
        root.FindAll(TreeScope_Descendants, &condition)
            .map_err(|e| WindowsMcpError::TreeError(format!("FindAll: {e}")))?
    };

    let limit = criteria.limit.clamp(1, MAX_FIND_LIMIT);
    let count = unsafe { elements.Length().unwrap_or(0) };

    let mut results = Vec::with_capacity(count.min(limit as i32) as usize);
    for i in 0..count {
        if results.len() >= limit {
            break;
        }
        if let Ok(elem) = unsafe { elements.GetElement(i) } {
            let info = unsafe { read_element_info(&elem) };

            // Apply name substring filter (case-insensitive) client-side
            // since UIA PropertyCondition for Name is exact match only.
            if let Some(ref name_filter) = criteria.name {
                if !info.name.to_lowercase().contains(&name_filter.to_lowercase()) {
                    continue;
                }
            }

            results.push(info);
        }
    }

    Ok(results)
}

/// Build a UIA condition from [`FindCriteria`].
///
/// - If `automation_id` is set, creates a PropertyCondition on AutomationId.
/// - If `control_type` is set, creates a PropertyCondition on ControlType name.
/// - Otherwise, uses `CreateTrueCondition` (match all).
///
/// Name filtering is done client-side because UIA PropertyCondition on Name
/// only supports exact match, not substring.
unsafe fn build_find_condition(
    uia: &IUIAutomation,
    criteria: &FindCriteria,
) -> Result<IUIAutomationCondition, WindowsMcpError> {
    let mut conditions: Vec<IUIAutomationCondition> = Vec::new();

    // AutomationId -- exact match
    if let Some(ref aid) = criteria.automation_id {
        let variant = windows::core::VARIANT::from(windows::core::BSTR::from(aid.as_str()));
        let cond = uia
            .CreatePropertyCondition(UIA_AutomationIdPropertyId, &variant)
            .map_err(|e| WindowsMcpError::TreeError(format!("CreatePropertyCondition(AutomationId): {e}")))?;
        conditions.push(cond.cast::<IUIAutomationCondition>().map_err(|e| WindowsMcpError::TreeError(format!("cast AutomationId condition: {e}")))?);
    }

    // ControlType -- convert name to ID, then exact match
    if let Some(ref ct_name) = criteria.control_type {
        if let Some(ct_id) = control_type_id_from_name(ct_name) {
            let variant = windows::core::VARIANT::from(ct_id);
            let cond = uia
                .CreatePropertyCondition(UIA_ControlTypePropertyId, &variant)
                .map_err(|e| WindowsMcpError::TreeError(format!("CreatePropertyCondition(ControlType): {e}")))?;
            conditions.push(cond.cast::<IUIAutomationCondition>().map_err(|e| WindowsMcpError::TreeError(format!("cast ControlType condition: {e}")))?);
        }
    }

    // Name -- UIA only supports exact match, so we use TrueCondition and
    // filter client-side in find_elements().  But if name is the ONLY
    // criterion, we still need at least a TrueCondition.

    match conditions.len() {
        0 => {
            let cond = uia
                .CreateTrueCondition()
                .map_err(|e| WindowsMcpError::TreeError(format!("CreateTrueCondition: {e}")))?;
            Ok(cond.cast::<IUIAutomationCondition>().map_err(|e| WindowsMcpError::TreeError(format!("cast TrueCondition: {e}")))?)
        }
        1 => Ok(conditions.remove(0)),
        _ => {
            // Chain with AND
            let mut combined = conditions[0].clone();
            for cond in &conditions[1..] {
                combined = uia
                    .CreateAndCondition(&combined, cond)
                    .map_err(|e| WindowsMcpError::TreeError(format!("CreateAndCondition: {e}")))?
                    .cast::<IUIAutomationCondition>()
                    .map_err(|e| WindowsMcpError::TreeError(format!("cast AndCondition: {e}")))?;
            }
            Ok(combined)
        }
    }
}

/// Map a control type name (e.g. "Button") to its UIA_*ControlTypeId integer.
///
/// Returns `None` for unrecognised names.
fn control_type_id_from_name(name: &str) -> Option<i32> {
    use windows::Win32::UI::Accessibility::*;
    match name {
        "AppBar" => Some(UIA_AppBarControlTypeId.0),
        "Button" => Some(UIA_ButtonControlTypeId.0),
        "Calendar" => Some(UIA_CalendarControlTypeId.0),
        "CheckBox" => Some(UIA_CheckBoxControlTypeId.0),
        "ComboBox" => Some(UIA_ComboBoxControlTypeId.0),
        "Custom" => Some(UIA_CustomControlTypeId.0),
        "DataGrid" => Some(UIA_DataGridControlTypeId.0),
        "DataItem" => Some(UIA_DataItemControlTypeId.0),
        "Document" => Some(UIA_DocumentControlTypeId.0),
        "Edit" => Some(UIA_EditControlTypeId.0),
        "Group" => Some(UIA_GroupControlTypeId.0),
        "Header" => Some(UIA_HeaderControlTypeId.0),
        "HeaderItem" => Some(UIA_HeaderItemControlTypeId.0),
        "Hyperlink" => Some(UIA_HyperlinkControlTypeId.0),
        "Image" => Some(UIA_ImageControlTypeId.0),
        "List" => Some(UIA_ListControlTypeId.0),
        "ListItem" => Some(UIA_ListItemControlTypeId.0),
        "MenuBar" => Some(UIA_MenuBarControlTypeId.0),
        "Menu" => Some(UIA_MenuControlTypeId.0),
        "MenuItem" => Some(UIA_MenuItemControlTypeId.0),
        "Pane" => Some(UIA_PaneControlTypeId.0),
        "ProgressBar" => Some(UIA_ProgressBarControlTypeId.0),
        "RadioButton" => Some(UIA_RadioButtonControlTypeId.0),
        "ScrollBar" => Some(UIA_ScrollBarControlTypeId.0),
        "SemanticZoom" => Some(UIA_SemanticZoomControlTypeId.0),
        "Separator" => Some(UIA_SeparatorControlTypeId.0),
        "Slider" => Some(UIA_SliderControlTypeId.0),
        "Spinner" => Some(UIA_SpinnerControlTypeId.0),
        "SplitButton" => Some(UIA_SplitButtonControlTypeId.0),
        "StatusBar" => Some(UIA_StatusBarControlTypeId.0),
        "Tab" => Some(UIA_TabControlTypeId.0),
        "TabItem" => Some(UIA_TabItemControlTypeId.0),
        "Table" => Some(UIA_TableControlTypeId.0),
        "Text" => Some(UIA_TextControlTypeId.0),
        "Thumb" => Some(UIA_ThumbControlTypeId.0),
        "TitleBar" => Some(UIA_TitleBarControlTypeId.0),
        "ToolBar" => Some(UIA_ToolBarControlTypeId.0),
        "ToolTip" => Some(UIA_ToolTipControlTypeId.0),
        "Tree" => Some(UIA_TreeControlTypeId.0),
        "TreeItem" => Some(UIA_TreeItemControlTypeId.0),
        "Window" => Some(UIA_WindowControlTypeId.0),
        _ => None,
    }
}

/// Query primary and virtual screen dimensions.
///
/// Uses `GetSystemMetrics` (not cached -- resolution can change at runtime).
pub fn get_screen_metrics() -> Result<ScreenMetrics, WindowsMcpError> {
    let (pw, ph, vw, vh) = unsafe {
        (
            GetSystemMetrics(SM_CXSCREEN),
            GetSystemMetrics(SM_CYSCREEN),
            GetSystemMetrics(SM_CXVIRTUALSCREEN),
            GetSystemMetrics(SM_CYVIRTUALSCREEN),
        )
    };

    if pw <= 0 || ph <= 0 {
        return Err(WindowsMcpError::ComError(
            "GetSystemMetrics returned non-positive primary screen dimensions".into(),
        ));
    }

    Ok(ScreenMetrics {
        primary_width: pw,
        primary_height: ph,
        virtual_width: if vw > 0 { vw } else { pw },
        virtual_height: if vh > 0 { vh } else { ph },
    })
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_find_criteria_default() {
        let c = FindCriteria::default();
        assert!(c.name.is_none());
        assert!(c.control_type.is_none());
        assert!(c.automation_id.is_none());
        assert!(c.window_handle.is_none());
        assert_eq!(c.limit, 0);
    }

    #[test]
    fn test_screen_metrics_serialization() {
        let m = ScreenMetrics {
            primary_width: 1920,
            primary_height: 1080,
            virtual_width: 3840,
            virtual_height: 1080,
        };
        let json = serde_json::to_string(&m).unwrap();
        assert!(json.contains("1920"));
        assert!(json.contains("3840"));
    }

    #[test]
    fn test_element_info_serialization() {
        let info = ElementInfo {
            name: "OK".into(),
            automation_id: "btn_ok".into(),
            control_type: "Button".into(),
            localized_control_type: "button".into(),
            class_name: "Button".into(),
            bounding_rect: [10.0, 20.0, 110.0, 50.0],
            is_enabled: true,
            is_offscreen: false,
            has_keyboard_focus: false,
            supported_patterns: vec!["InvokePattern".into()],
        };
        let json = serde_json::to_string(&info).unwrap();
        assert!(json.contains("\"name\":\"OK\""));
        assert!(json.contains("InvokePattern"));
    }

    #[test]
    fn test_control_type_id_from_name_known() {
        assert!(control_type_id_from_name("Button").is_some());
        assert!(control_type_id_from_name("Edit").is_some());
        assert!(control_type_id_from_name("Window").is_some());
    }

    #[test]
    fn test_control_type_id_from_name_unknown() {
        assert!(control_type_id_from_name("NonExistent").is_none());
        assert!(control_type_id_from_name("").is_none());
    }

    #[test]
    fn test_max_find_limit_clamp() {
        assert_eq!(200_usize.clamp(1, MAX_FIND_LIMIT), MAX_FIND_LIMIT);
        assert_eq!(0_usize.clamp(1, MAX_FIND_LIMIT), 1);
        assert_eq!(50_usize.clamp(1, MAX_FIND_LIMIT), 50);
    }
}
