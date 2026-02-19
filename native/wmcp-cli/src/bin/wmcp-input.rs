//! Standalone CLI tool for sending input events.

use clap::{Parser, Subcommand};

#[derive(Parser)]
#[command(name = "wmcp-input", about = "Send keyboard/mouse input via Win32 SendInput")]
struct Args {
    #[command(subcommand)]
    command: Command,
}

#[derive(Subcommand)]
enum Command {
    /// Type Unicode text
    Text {
        /// The text to type
        text: String,
    },
    /// Click at screen coordinates
    Click {
        /// X coordinate
        x: i32,
        /// Y coordinate
        y: i32,
        /// Button: left, right, middle
        #[arg(short, long, default_value = "left")]
        button: String,
    },
    /// Press a virtual key code
    Key {
        /// Virtual key code (hex, e.g. 0x0D for Enter)
        #[arg(value_parser = parse_hex_or_dec)]
        vk_code: u16,
    },
    /// Move cursor to coordinates
    Move {
        /// X coordinate
        x: i32,
        /// Y coordinate
        y: i32,
    },
    /// Send a hotkey combination
    Hotkey {
        /// Virtual key codes (hex, e.g. 0x11 0x43 for Ctrl+C)
        #[arg(value_parser = parse_hex_or_dec)]
        vk_codes: Vec<u16>,
    },
}

fn parse_hex_or_dec(s: &str) -> Result<u16, String> {
    if let Some(hex) = s.strip_prefix("0x").or_else(|| s.strip_prefix("0X")) {
        u16::from_str_radix(hex, 16).map_err(|e| e.to_string())
    } else {
        s.parse::<u16>().map_err(|e| e.to_string())
    }
}

fn main() {
    let args = Args::parse();

    match args.command {
        Command::Text { text } => {
            let count = wmcp_core::input::send_text_raw(&text);
            println!("Sent {count} events for {} chars", text.len());
        }
        Command::Click { x, y, button } => {
            let count = wmcp_core::input::send_click_raw(x, y, &button);
            println!("Sent {count} events (click {button} at {x},{y})");
        }
        Command::Key { vk_code } => {
            wmcp_core::input::send_key_raw(vk_code, false);
            wmcp_core::input::send_key_raw(vk_code, true);
            println!("Sent key 0x{vk_code:04X}");
        }
        Command::Move { x, y } => {
            let count = wmcp_core::input::send_mouse_move_raw(x, y);
            println!("Moved cursor to {x},{y} ({count} events)");
        }
        Command::Hotkey { vk_codes } => {
            let count = wmcp_core::input::send_hotkey_raw(&vk_codes);
            let hex: Vec<String> = vk_codes.iter().map(|v| format!("0x{v:04X}")).collect();
            println!("Sent hotkey [{}] ({count} events)", hex.join("+"));
        }
    }
}
