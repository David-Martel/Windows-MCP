//! JSON-RPC IPC worker for COM-isolated operations.
//!
//! Reads line-delimited JSON requests from stdin, dispatches to wmcp_core,
//! writes JSON responses to stdout.

use std::io::{self, BufRead, Write};

use clap::Parser;
use serde::{Deserialize, Serialize};

#[derive(Parser)]
#[command(name = "wmcp-worker", about = "Windows-MCP IPC worker process")]
struct Args {
    /// Enable verbose logging to stderr
    #[arg(short, long)]
    verbose: bool,
}

#[derive(Deserialize)]
struct Request {
    id: u64,
    method: String,
    #[serde(default)]
    params: serde_json::Value,
}

#[derive(Serialize)]
struct Response {
    id: u64,
    #[serde(skip_serializing_if = "Option::is_none")]
    result: Option<serde_json::Value>,
    #[serde(skip_serializing_if = "Option::is_none")]
    error: Option<String>,
}

fn dispatch(method: &str, params: &serde_json::Value) -> Result<serde_json::Value, String> {
    match method {
        "system_info" => {
            let snapshot = wmcp_core::system_info::collect_system_info()
                .map_err(|e| e.to_string())?;
            serde_json::to_value(snapshot).map_err(|e| e.to_string())
        }
        "capture_tree" => {
            let handles: Vec<isize> = params
                .get("handles")
                .and_then(|v| serde_json::from_value(v.clone()).ok())
                .unwrap_or_default();
            let max_depth: usize = params
                .get("max_depth")
                .and_then(|v| v.as_u64())
                .unwrap_or(50) as usize;
            let snapshots = wmcp_core::tree::capture_tree_raw(&handles, max_depth);
            serde_json::to_value(snapshots).map_err(|e| e.to_string())
        }
        "send_text" => {
            let text = params
                .get("text")
                .and_then(|v| v.as_str())
                .unwrap_or("");
            let count = wmcp_core::input::send_text_raw(text);
            Ok(serde_json::Value::from(count))
        }
        "send_click" => {
            let x = params.get("x").and_then(|v| v.as_i64()).unwrap_or(0) as i32;
            let y = params.get("y").and_then(|v| v.as_i64()).unwrap_or(0) as i32;
            let button = params.get("button").and_then(|v| v.as_str()).unwrap_or("left");
            let count = wmcp_core::input::send_click_raw(x, y, button);
            Ok(serde_json::Value::from(count))
        }
        "send_key" => {
            let vk = params.get("vk_code").and_then(|v| v.as_u64()).unwrap_or(0) as u16;
            let key_up = params.get("key_up").and_then(|v| v.as_bool()).unwrap_or(false);
            let count = wmcp_core::input::send_key_raw(vk, key_up);
            Ok(serde_json::Value::from(count))
        }
        "send_hotkey" => {
            let vk_codes: Vec<u16> = params
                .get("vk_codes")
                .and_then(|v| serde_json::from_value(v.clone()).ok())
                .unwrap_or_default();
            let count = wmcp_core::input::send_hotkey_raw(&vk_codes);
            Ok(serde_json::Value::from(count))
        }
        "ping" => Ok(serde_json::Value::String("pong".to_owned())),
        _ => Err(format!("unknown method: {method}")),
    }
}

fn main() {
    let args = Args::parse();
    let stdin = io::stdin();
    let mut stdout = io::stdout();

    if args.verbose {
        eprintln!("wmcp-worker: ready");
    }

    for line in stdin.lock().lines() {
        let line = match line {
            Ok(l) => l,
            Err(e) => {
                if args.verbose {
                    eprintln!("wmcp-worker: stdin read error: {e}");
                }
                break;
            }
        };

        if line.trim().is_empty() {
            continue;
        }

        let req: Request = match serde_json::from_str(&line) {
            Ok(r) => r,
            Err(e) => {
                let resp = Response {
                    id: 0,
                    result: None,
                    error: Some(format!("invalid JSON: {e}")),
                };
                let _ = writeln!(stdout, "{}", serde_json::to_string(&resp).unwrap());
                let _ = stdout.flush();
                continue;
            }
        };

        let resp = match dispatch(&req.method, &req.params) {
            Ok(result) => Response {
                id: req.id,
                result: Some(result),
                error: None,
            },
            Err(error) => Response {
                id: req.id,
                result: None,
                error: Some(error),
            },
        };

        let _ = writeln!(stdout, "{}", serde_json::to_string(&resp).unwrap());
        let _ = stdout.flush();
    }
}
