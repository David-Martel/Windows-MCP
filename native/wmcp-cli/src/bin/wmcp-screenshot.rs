//! Standalone CLI tool for capturing screenshots.
//!
//! Placeholder -- DXGI screenshot module will be added in Phase 3.

use clap::Parser;

#[derive(Parser)]
#[command(name = "wmcp-screenshot", about = "Capture screenshot via DXGI (placeholder)")]
struct Args {
    /// Output file path
    #[arg(short, long, default_value = "screenshot.png")]
    output: String,

    /// Monitor index (0 = primary)
    #[arg(long, default_value = "0")]
    monitor: u32,
}

fn main() {
    let args = Args::parse();
    eprintln!(
        "wmcp-screenshot: DXGI capture not yet implemented. \
         Would save monitor {} to '{}'",
        args.monitor, args.output
    );
    std::process::exit(1);
}
