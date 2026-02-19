//! System information via the `sysinfo` crate.
//!
//! Replaces PowerShell subprocess calls (200-500ms each) with a single
//! in-process Rust call that takes ~1-5ms on first use and <1ms on
//! subsequent calls.
//!
//! # Thread safety
//!
//! `sysinfo::System` is wrapped in `parking_lot::Mutex` + `OnceLock` for
//! safe concurrent access.

use std::sync::OnceLock;

use parking_lot::Mutex;
use serde::Serialize;
use sysinfo::{CpuRefreshKind, Disks, MemoryRefreshKind, RefreshKind, System};

use crate::errors::WindowsMcpError;

// ---------------------------------------------------------------------------
// Singleton
// ---------------------------------------------------------------------------

static SYSTEM: OnceLock<Mutex<System>> = OnceLock::new();

fn get_system() -> &'static Mutex<System> {
    SYSTEM.get_or_init(|| {
        Mutex::new(System::new_with_specifics(
            RefreshKind::nothing()
                .with_cpu(CpuRefreshKind::everything())
                .with_memory(MemoryRefreshKind::everything()),
        ))
    })
}

// ---------------------------------------------------------------------------
// Data transfer objects
// ---------------------------------------------------------------------------

/// Owned snapshot of system state -- fully `Send` and serializable.
#[derive(Debug, Clone, Serialize)]
pub struct SystemSnapshot {
    pub os_name: String,
    pub os_version: String,
    pub hostname: String,
    pub cpu_count: usize,
    pub cpu_usage: Vec<f32>,
    pub total_memory_bytes: u64,
    pub used_memory_bytes: u64,
    pub disks: Vec<DiskSnapshot>,
}

/// Owned snapshot of a single disk.
#[derive(Debug, Clone, Serialize)]
pub struct DiskSnapshot {
    pub name: String,
    pub mount_point: String,
    pub total_bytes: u64,
    pub available_bytes: u64,
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/// Collect system information and return an owned snapshot.
///
/// This function is blocking (holds the sysinfo mutex).  PyO3 callers
/// should wrap it in `py.allow_threads()`.
pub fn collect_system_info() -> Result<SystemSnapshot, WindowsMcpError> {
    let mutex = get_system();
    let mut sys = mutex.lock();

    sys.refresh_cpu_usage();
    sys.refresh_memory();

    let cpu_usage: Vec<f32> = sys.cpus().iter().map(|c| c.cpu_usage()).collect();
    let cpu_count = sys.cpus().len();

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
}
