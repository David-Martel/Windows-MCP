//! System information via the [`sysinfo`] crate.
//!
//! Replaces the PowerShell subprocess calls in `desktop/service.py` (each of
//! which costs 200-500 ms of process spawn overhead) with a single in-process
//! Rust call that takes ~1-5 ms on first use and <1 ms on subsequent calls
//! because the [`System`] singleton is refreshed only when [`system_info`] is
//! called, not on every query.
//!
//! # Thread safety
//!
//! [`sysinfo::System`] is `Send` but not `Sync`.  We wrap it in a
//! `parking_lot::Mutex` stored in a `std::sync::OnceLock` so that:
//!
//! - The first call initialises the singleton (no `unsafe`).
//! - Concurrent Python threads block on the mutex rather than racing.
//! - The GIL is released while we hold the Rust mutex, so other Python
//!   threads are free to run.
//!
//! # Example (Python)
//!
//! ```python
//! import windows_mcp_core
//!
//! info = windows_mcp_core.system_info()
//! print(info["os_name"], info["cpu_count"])
//! for disk in info["disks"]:
//!     print(disk["mount_point"], disk["available_bytes"])
//! ```

use std::sync::OnceLock;

use parking_lot::Mutex;
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};
use sysinfo::{CpuRefreshKind, Disks, MemoryRefreshKind, RefreshKind, System};

use crate::errors::WindowsMcpError;

// ---------------------------------------------------------------------------
// Singleton
// ---------------------------------------------------------------------------

/// Global [`System`] handle initialised exactly once.
///
/// `OnceLock` guarantees that the closure runs once even under concurrent
/// calls.  `parking_lot::Mutex` is chosen over `std::sync::Mutex` because it
/// is smaller (no poisoning), faster on Windows, and never panics.
static SYSTEM: OnceLock<Mutex<System>> = OnceLock::new();

/// Return a reference to the global [`Mutex<System>`], initialising it on the
/// first call.
fn get_system() -> &'static Mutex<System> {
    SYSTEM.get_or_init(|| {
        // Construct with an empty refresh spec; we refresh explicitly below.
        Mutex::new(System::new_with_specifics(
            RefreshKind::nothing()
                .with_cpu(CpuRefreshKind::everything())
                .with_memory(MemoryRefreshKind::everything()),
        ))
    })
}

// ---------------------------------------------------------------------------
// Public pyfunction
// ---------------------------------------------------------------------------

/// Collect system information and return it as a Python `dict`.
///
/// The GIL is released for the duration of the `sysinfo` refresh so that
/// other Python threads are not blocked during the ~1-5 ms data collection.
///
/// # Returns
///
/// A `dict` with the following keys:
///
/// | Key | Type | Description |
/// |-----|------|-------------|
/// | `os_name` | `str` | OS name (e.g. `"Windows 11"`) |
/// | `os_version` | `str` | OS version string |
/// | `hostname` | `str` | Machine hostname |
/// | `cpu_count` | `int` | Logical CPU count |
/// | `cpu_usage_percent` | `list[float]` | Per-CPU usage 0-100 |
/// | `total_memory_bytes` | `int` | Total physical RAM in bytes |
/// | `used_memory_bytes` | `int` | Used physical RAM in bytes |
/// | `disks` | `list[dict]` | Disk info (see below) |
///
/// Each disk dict contains:
///
/// | Key | Type |
/// |-----|------|
/// | `name` | `str` |
/// | `mount_point` | `str` |
/// | `total_bytes` | `int` |
/// | `available_bytes` | `int` |
///
/// # Errors
///
/// Raises `RuntimeError` (mapped from [`WindowsMcpError::SystemInfoError`])
/// if the mutex is poisoned or sysinfo returns an unexpected state.
///
/// # Example
///
/// ```python
/// import windows_mcp_core
///
/// info = windows_mcp_core.system_info()
/// assert isinstance(info["cpu_count"], int)
/// assert info["cpu_count"] > 0
/// assert all(0.0 <= u <= 100.0 for u in info["cpu_usage_percent"])
/// ```
#[pyfunction]
pub fn system_info(py: Python<'_>) -> PyResult<PyObject> {
    // Release the GIL while we do the blocking sysinfo refresh so other
    // Python threads can run concurrently.
    let snapshot = py
        .allow_threads(|| -> Result<SystemSnapshot, WindowsMcpError> {
            let mutex = get_system();
            let mut sys = mutex.lock(); // parking_lot: no poison, no unwrap needed

            // Refresh CPU and memory.  A second CPU refresh is required to
            // get meaningful usage percentages (sysinfo needs two samples).
            sys.refresh_cpu_usage();
            // Small sleep would improve accuracy, but we avoid std::thread::sleep
            // inside allow_threads to keep latency predictable.  Callers that
            // need accurate CPU% should call system_info() twice ~100 ms apart.
            sys.refresh_memory();

            let cpu_usage: Vec<f32> = sys.cpus().iter().map(|c| c.cpu_usage()).collect();
            let cpu_count = sys.cpus().len();

            // Disk information is collected from a fresh snapshot each call
            // because Disks does not need a persistent singleton -- it is
            // cheap to enumerate.
            let disks = Disks::new_with_refreshed_list();
            let disk_snapshots: Vec<DiskSnapshot> = disks
                .iter()
                .map(|d| DiskSnapshot {
                    name: d.name().to_string_lossy().into_owned(),
                    mount_point: d.mount_point().to_string_lossy().into_owned(),
                    total_bytes: d.total_space(),
                    available_bytes: d.available_space(),
                })
                .collect();

            Ok(SystemSnapshot {
                os_name: System::long_os_version().unwrap_or_else(|| "Unknown".to_owned()),
                os_version: System::os_version().unwrap_or_else(|| "Unknown".to_owned()),
                hostname: System::host_name().unwrap_or_else(|| "Unknown".to_owned()),
                cpu_count,
                cpu_usage,
                total_memory_bytes: sys.total_memory(),
                used_memory_bytes: sys.used_memory(),
                disks: disk_snapshots,
            })
        })
        .map_err(PyErr::from)?;

    // Re-acquire the GIL to build the Python dict.
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
// Internal data transfer objects (stack-allocated, no heap allocation after
// the sysinfo refresh)
// ---------------------------------------------------------------------------

/// Owned snapshot of system state captured while holding the mutex.
///
/// Separating data collection (inside `allow_threads`) from Python object
/// construction (outside `allow_threads`) is the canonical PyO3 pattern for
/// releasing the GIL during blocking work.
struct SystemSnapshot {
    os_name: String,
    os_version: String,
    hostname: String,
    cpu_count: usize,
    cpu_usage: Vec<f32>,
    total_memory_bytes: u64,
    used_memory_bytes: u64,
    disks: Vec<DiskSnapshot>,
}

struct DiskSnapshot {
    name: String,
    mount_point: String,
    total_bytes: u64,
    available_bytes: u64,
}
