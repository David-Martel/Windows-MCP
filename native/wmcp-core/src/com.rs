//! COM apartment RAII guard.
//!
//! [`COMGuard`] wraps `CoInitializeEx` / `CoUninitialize` in an RAII pattern
//! so that COM apartments are correctly initialised and cleaned up, even on
//! panic or early return.
//!
//! The `PhantomData<*const ()>` field enforces `!Send` + `!Sync` at compile
//! time, preventing the guard from being moved across thread boundaries.

use crate::errors::WindowsMcpError;
use log;
use windows::Win32::System::Com::{CoInitializeEx, CoUninitialize, COINIT_MULTITHREADED};

/// RAII wrapper that calls `CoUninitialize` on `Drop` when appropriate.
///
/// Instantiate **once per thread** via [`COMGuard::init`].  The guard tracks
/// whether `CoInitializeEx` actually succeeded (vs. `RPC_E_CHANGED_MODE`)
/// and only calls `CoUninitialize` when a balancing call is required per MSDN.
#[must_use = "COMGuard must be kept alive for the duration of COM usage"]
pub struct COMGuard {
    should_uninit: bool,
    _not_send: std::marker::PhantomData<*const ()>,
}

impl COMGuard {
    /// Initialise (or join) the thread's MTA COM apartment.
    ///
    /// Returns `Ok(COMGuard)` for `S_OK`, `S_FALSE`, and
    /// `RPC_E_CHANGED_MODE` (thread has STA; COM is usable but we must
    /// NOT call `CoUninitialize` since we did not successfully initialise).
    pub fn init() -> Result<Self, WindowsMcpError> {
        let hr = unsafe { CoInitializeEx(None, COINIT_MULTITHREADED) };

        let hresult_value = hr.0 as u32;
        match hresult_value {
            // S_OK -- newly initialised.
            0x0 => Ok(COMGuard {
                should_uninit: true,
                _not_send: std::marker::PhantomData,
            }),
            // S_FALSE -- already initialised on this thread.
            0x1 => Ok(COMGuard {
                should_uninit: true,
                _not_send: std::marker::PhantomData,
            }),
            // RPC_E_CHANGED_MODE -- thread already has STA.  COM is usable
            // but we requested MTA, so log a warning for diagnostics.
            0x80010106 => {
                log::warn!(
                    "CoInitializeEx: RPC_E_CHANGED_MODE -- thread already has STA apartment, \
                     using existing apartment instead of MTA"
                );
                Ok(COMGuard {
                    should_uninit: false,
                    _not_send: std::marker::PhantomData,
                })
            }
            _ => Err(WindowsMcpError::ComError(format!(
                "CoInitializeEx failed: HRESULT 0x{hresult_value:08X}"
            ))),
        }
    }
}

impl Drop for COMGuard {
    fn drop(&mut self) {
        if self.should_uninit {
            unsafe { CoUninitialize() };
        }
    }
}
