"""ctypes-based FFI wrapper for windows_mcp_ffi.dll.

Provides the same interface as the PyO3 extension but via C ABI,
useful when PyO3 won't build or hot-reload is needed.

Usage::

    from windows_mcp.native_ffi import NativeFFI

    ffi = NativeFFI()  # Loads windows_mcp_ffi.dll
    info = ffi.system_info()
    print(info["os_name"])
"""

import ctypes
import json
import logging
import os
from ctypes import (
    POINTER,
    c_char_p,
    c_int32,
    c_size_t,
    c_uint32,
    pointer,
)
from pathlib import Path

logger = logging.getLogger(__name__)

WMCP_OK = 0
WMCP_ERROR = -1


def _find_ffi_dll() -> Path | None:
    """Search for windows_mcp_ffi.dll in known locations."""
    candidates = [
        # Next to this source file
        Path(__file__).parent / "windows_mcp_ffi.dll",
        # Shared Cargo target
        Path("T:/RustCache/cargo-target/release/windows_mcp_ffi.dll"),
        # Local build
        Path("native/target/release/windows_mcp_ffi.dll"),
    ]

    # Also check env var
    env_path = os.environ.get("WMCP_FFI_DLL")
    if env_path:
        candidates.insert(0, Path(env_path))

    for p in candidates:
        if p.exists():
            return p
    return None


class NativeFFI:
    """ctypes wrapper around the windows_mcp_ffi C ABI DLL."""

    def __init__(self, dll_path: str | Path | None = None):
        if dll_path is None:
            found = _find_ffi_dll()
            if found is None:
                raise FileNotFoundError(
                    "windows_mcp_ffi.dll not found. Build with "
                    "`cargo build --release -p wmcp-ffi` or set WMCP_FFI_DLL env var."
                )
            dll_path = found

        self._dll = ctypes.CDLL(str(dll_path))
        self._setup_prototypes()
        logger.info("NativeFFI loaded from %s", dll_path)

    def _setup_prototypes(self):
        dll = self._dll

        # wmcp_last_error() -> *const c_char
        dll.wmcp_last_error.restype = c_char_p
        dll.wmcp_last_error.argtypes = []

        # wmcp_free_string(*mut c_char)
        dll.wmcp_free_string.restype = None
        dll.wmcp_free_string.argtypes = [c_char_p]

        # wmcp_system_info(*mut *mut c_char) -> i32
        dll.wmcp_system_info.restype = c_int32
        dll.wmcp_system_info.argtypes = [POINTER(c_char_p)]

        # wmcp_send_text(*const c_char, *mut u32) -> i32
        dll.wmcp_send_text.restype = c_int32
        dll.wmcp_send_text.argtypes = [c_char_p, POINTER(c_uint32)]

        # wmcp_send_click(i32, i32, i32) -> i32
        dll.wmcp_send_click.restype = c_int32
        dll.wmcp_send_click.argtypes = [c_int32, c_int32, c_int32]

        # wmcp_capture_tree(*const isize, usize, usize, *mut *mut c_char) -> i32
        dll.wmcp_capture_tree.restype = c_int32
        dll.wmcp_capture_tree.argtypes = [
            ctypes.POINTER(ctypes.c_ssize_t),
            c_size_t,
            c_size_t,
            POINTER(c_char_p),
        ]

    def _check_error(self, status: int, operation: str):
        if status != WMCP_OK:
            err = self._dll.wmcp_last_error()
            msg = err.decode("utf-8") if err else "unknown error"
            raise RuntimeError(f"{operation} failed: {msg}")

    def system_info(self) -> dict:
        """Collect system information, returns a dict."""
        out = c_char_p()
        status = self._dll.wmcp_system_info(pointer(out))
        self._check_error(status, "wmcp_system_info")
        try:
            return json.loads(out.value.decode("utf-8"))
        finally:
            self._dll.wmcp_free_string(out)

    def send_text(self, text: str) -> int:
        """Type text via SendInput, returns event count."""
        out_count = c_uint32()
        status = self._dll.wmcp_send_text(text.encode("utf-8"), pointer(out_count))
        self._check_error(status, "wmcp_send_text")
        return out_count.value

    def send_click(self, x: int, y: int, button: str = "left") -> None:
        """Click at screen coordinates."""
        button_int = {"left": 0, "right": 1, "middle": 2}.get(button, 0)
        status = self._dll.wmcp_send_click(x, y, button_int)
        self._check_error(status, "wmcp_send_click")

    def capture_tree(self, handles: list[int], max_depth: int = 50) -> list[dict]:
        """Capture UIA tree for window handles, returns list of dicts."""
        HandleArray = ctypes.c_ssize_t * len(handles)
        arr = HandleArray(*handles)
        out = c_char_p()
        status = self._dll.wmcp_capture_tree(arr, len(handles), max_depth, pointer(out))
        self._check_error(status, "wmcp_capture_tree")
        try:
            return json.loads(out.value.decode("utf-8"))
        finally:
            self._dll.wmcp_free_string(out)
