//! Standalone CLI tool for dumping the UIA accessibility tree as JSON.

use clap::Parser;

#[derive(Parser)]
#[command(name = "wmcp-tree", about = "Dump Windows UI Automation tree as JSON")]
struct Args {
    /// Window handle(s) to capture. If omitted, captures the foreground window.
    #[arg(long)]
    hwnd: Vec<isize>,

    /// Capture all visible windows
    #[arg(long)]
    all: bool,

    /// Maximum tree depth
    #[arg(long, default_value = "50")]
    max_depth: usize,

    /// Compact JSON output (no pretty-printing)
    #[arg(long)]
    compact: bool,
}

fn get_foreground_hwnd() -> isize {
    use windows::Win32::UI::WindowsAndMessaging::GetForegroundWindow;
    let hwnd = unsafe { GetForegroundWindow() };
    hwnd.0 as isize
}

fn main() {
    let args = Args::parse();

    let handles = if args.all {
        wmcp_core::window::enumerate_visible_windows().unwrap_or_else(|e| {
            eprintln!("Failed to enumerate windows: {e}");
            vec![get_foreground_hwnd()]
        })
    } else if args.hwnd.is_empty() {
        vec![get_foreground_hwnd()]
    } else {
        args.hwnd
    };

    let snapshots = wmcp_core::tree::capture_tree_raw(&handles, args.max_depth);

    let json = if args.compact {
        serde_json::to_string(&snapshots).unwrap()
    } else {
        serde_json::to_string_pretty(&snapshots).unwrap()
    };

    println!("{json}");
}
