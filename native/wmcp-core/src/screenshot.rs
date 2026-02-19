//! Desktop screenshot capture via DXGI Output Duplication.
//!
//! Captures the desktop without a window handle (headless), supporting
//! multi-monitor setups via a monitor index.  Falls back to a GDI
//! `BitBlt` capture when DXGI Output Duplication is unavailable (e.g.
//! inside a Remote Desktop session or on Windows Server SKUs without a
//! hardware-accelerated session).
//!
//! # DXGI Output Duplication overview
//!
//! 1. Create a `D3D11Device` with a hardware or WARP adapter.
//! 2. Walk `IDXGIFactory1 -> IDXGIAdapter -> IDXGIOutput -> IDXGIOutput1`.
//! 3. Call `IDXGIOutput1::DuplicateOutput` to obtain an
//!    `IDXGIOutputDuplication` interface.
//! 4. `AcquireNextFrame` returns a `IDXGISurface` backed by a GPU texture.
//! 5. Create a CPU-readable staging texture (`D3D11_USAGE_STAGING`), copy
//!    the desktop frame into it, then map it with `D3D11_MAP_READ` to
//!    obtain a `*const u8` pointer to BGRA pixels.
//! 6. Release the frame and duplicate interface after reading.
//!
//! # Thread safety
//!
//! All DXGI / D3D11 interfaces are COM objects.  This module creates them
//! fresh on every call -- there is no shared global state.  Each call must
//! be made from a thread with a valid COM apartment (call [`crate::com::COMGuard::init`]
//! before invoking these functions from a new thread).
//!
//! # Examples
//!
//! ```no_run
//! use wmcp_core::screenshot::{capture_raw, capture_png};
//!
//! // Capture primary monitor as raw BGRA bytes
//! let frame = capture_raw(0).expect("capture failed");
//! println!("{}x{} {} bytes", frame.width, frame.height, frame.data.len());
//!
//! // Capture primary monitor as a PNG
//! let png_bytes = capture_png(0).expect("PNG encode failed");
//! std::fs::write("screenshot.png", &png_bytes).unwrap();
//! ```

use windows::Win32::Foundation::RECT;
use windows::Win32::Graphics::Direct3D::D3D_DRIVER_TYPE_HARDWARE;
use windows::Win32::Graphics::Direct3D11::{
    D3D11CreateDevice, ID3D11Device, ID3D11DeviceContext, ID3D11Texture2D,
    D3D11_BIND_FLAG, D3D11_CPU_ACCESS_READ, D3D11_MAP_READ, D3D11_SDK_VERSION,
    D3D11_TEXTURE2D_DESC, D3D11_USAGE_STAGING,
};
use windows::Win32::Graphics::Dxgi::Common::{DXGI_FORMAT_B8G8R8A8_UNORM, DXGI_SAMPLE_DESC};
use windows::Win32::Graphics::Dxgi::{
    CreateDXGIFactory1, IDXGIAdapter, IDXGIFactory1, IDXGIOutput, IDXGIOutput1,
    IDXGIOutputDuplication, IDXGIResource, DXGI_OUTDUPL_FRAME_INFO, DXGI_OUTPUT_DESC,
};
use windows::core::Interface;

use crate::errors::WindowsMcpError;

// ---------------------------------------------------------------------------
// GDI fallback imports
// ---------------------------------------------------------------------------
use windows::Win32::Foundation::HWND;
use windows::Win32::Graphics::Gdi::{
    BitBlt, CreateCompatibleBitmap, CreateCompatibleDC, DeleteDC, DeleteObject, GetDC,
    GetDIBits, ReleaseDC, SelectObject, BITMAPINFO, BITMAPINFOHEADER, BI_RGB, DIB_RGB_COLORS,
    SRCCOPY,
};
use windows::Win32::UI::WindowsAndMessaging::{GetSystemMetrics, SM_CXSCREEN, SM_CYSCREEN};

// ---------------------------------------------------------------------------
// Public data types
// ---------------------------------------------------------------------------

/// Raw BGRA pixel data for a single monitor frame.
///
/// Pixels are stored row-major, left-to-right, top-to-bottom.
/// Each pixel is 4 bytes in BGRA order (matching the DXGI
/// `DXGI_FORMAT_B8G8R8A8_UNORM` layout).
#[derive(Debug, Clone)]
pub struct ScreenshotData {
    /// Width of the captured frame in pixels.
    pub width: u32,
    /// Height of the captured frame in pixels.
    pub height: u32,
    /// Raw pixel bytes in BGRA order; length == `width * height * 4`.
    pub data: Vec<u8>,
}

// ---------------------------------------------------------------------------
// Internal DXGI capture helpers
// ---------------------------------------------------------------------------

/// Create a `D3D11Device` and its immediate context.
///
/// Tries hardware first; falls back to the WARP software renderer on
/// failure.  The returned device is suitable for creating staging
/// textures to copy DXGI duplication frames into CPU-accessible memory.
fn create_d3d11_device() -> Result<(ID3D11Device, ID3D11DeviceContext), WindowsMcpError> {
    let mut device: Option<ID3D11Device> = None;
    let mut context: Option<ID3D11DeviceContext> = None;
    let mut feature_level =
        windows::Win32::Graphics::Direct3D::D3D_FEATURE_LEVEL_9_1;

    // Try hardware adapter first.
    let hr = unsafe {
        D3D11CreateDevice(
            None,
            D3D_DRIVER_TYPE_HARDWARE,
            None,
            windows::Win32::Graphics::Direct3D11::D3D11_CREATE_DEVICE_FLAG(0),
            None,
            D3D11_SDK_VERSION,
            Some(&mut device),
            Some(&mut feature_level),
            Some(&mut context),
        )
    };

    let (device, context) = if hr.is_ok() {
        (
            device.ok_or_else(|| {
                WindowsMcpError::ScreenshotError(
                    "D3D11CreateDevice succeeded but returned null device".into(),
                )
            })?,
            context.ok_or_else(|| {
                WindowsMcpError::ScreenshotError(
                    "D3D11CreateDevice succeeded but returned null context".into(),
                )
            })?,
        )
    } else {
        // Try WARP software renderer as fallback.
        let mut device2: Option<ID3D11Device> = None;
        let mut context2: Option<ID3D11DeviceContext> = None;

        unsafe {
            D3D11CreateDevice(
                None,
                windows::Win32::Graphics::Direct3D::D3D_DRIVER_TYPE_WARP,
                None,
                windows::Win32::Graphics::Direct3D11::D3D11_CREATE_DEVICE_FLAG(0),
                None,
                D3D11_SDK_VERSION,
                Some(&mut device2),
                Some(&mut feature_level),
                Some(&mut context2),
            )
            .map_err(|e| {
                WindowsMcpError::ScreenshotError(format!(
                    "D3D11CreateDevice (WARP fallback) failed: {e}"
                ))
            })?;
        }

        (
            device2.ok_or_else(|| {
                WindowsMcpError::ScreenshotError(
                    "D3D11CreateDevice (WARP) returned null device".into(),
                )
            })?,
            context2.ok_or_else(|| {
                WindowsMcpError::ScreenshotError(
                    "D3D11CreateDevice (WARP) returned null context".into(),
                )
            })?,
        )
    };

    Ok((device, context))
}

/// Enumerate DXGI outputs (monitors) and return the `IDXGIOutput1` for
/// `monitor_index`, plus the first adapter that owns it.
fn get_dxgi_output(
    monitor_index: u32,
) -> Result<(IDXGIAdapter, IDXGIOutput1, DXGI_OUTPUT_DESC), WindowsMcpError> {
    let factory: IDXGIFactory1 = unsafe {
        CreateDXGIFactory1().map_err(|e| {
            WindowsMcpError::ScreenshotError(format!("CreateDXGIFactory1 failed: {e}"))
        })?
    };

    let mut global_output_index: u32 = 0;

    // Walk adapters (graphics cards) in order.
    let mut adapter_index: u32 = 0;
    loop {
        let adapter: IDXGIAdapter = match unsafe { factory.EnumAdapters(adapter_index) } {
            Ok(a) => a,
            Err(_) => {
                // DXGI_ERROR_NOT_FOUND signals end of adapters.
                break;
            }
        };

        // Walk outputs (monitors) on this adapter.
        let mut output_index: u32 = 0;
        loop {
            let output: IDXGIOutput = match unsafe { adapter.EnumOutputs(output_index) } {
                Ok(o) => o,
                Err(_) => break, // end of outputs on this adapter
            };

            if global_output_index == monitor_index {
                let output1: IDXGIOutput1 = output.cast::<IDXGIOutput1>().map_err(|e| {
                    WindowsMcpError::ScreenshotError(format!(
                        "IDXGIOutput -> IDXGIOutput1 cast failed (monitor {monitor_index}): {e}"
                    ))
                })?;

                let desc = unsafe {
                    output1.GetDesc().map_err(|e| {
                        WindowsMcpError::ScreenshotError(format!(
                            "IDXGIOutput1::GetDesc failed: {e}"
                        ))
                    })?
                };

                return Ok((adapter, output1, desc));
            }

            global_output_index += 1;
            output_index += 1;
        }

        adapter_index += 1;
    }

    Err(WindowsMcpError::ScreenshotError(format!(
        "Monitor index {monitor_index} not found; system has {global_output_index} monitor(s)"
    )))
}

/// Capture one frame from `duplication`, copy it into a CPU-readable
/// staging texture, map it, and return the raw BGRA bytes.
///
/// The caller provides the D3D11 device and context so the staging
/// texture is created with the same device that owns the duplication.
fn read_frame(
    device: &ID3D11Device,
    context: &ID3D11DeviceContext,
    duplication: &IDXGIOutputDuplication,
    width: u32,
    height: u32,
) -> Result<Vec<u8>, WindowsMcpError> {
    // AcquireNextFrame blocks until a new frame is ready.
    // Timeout of 500ms is enough for a 60Hz display (~16ms between frames).
    let timeout_ms: u32 = 500;
    let mut frame_info = DXGI_OUTDUPL_FRAME_INFO::default();
    let mut desktop_resource: Option<IDXGIResource> = None;

    unsafe {
        duplication
            .AcquireNextFrame(timeout_ms, &mut frame_info, &mut desktop_resource)
            .map_err(|e| {
                WindowsMcpError::ScreenshotError(format!("AcquireNextFrame failed: {e}"))
            })?;
    }

    // The desktop resource is a `IDXGISurface` backed by a GPU texture.
    // We must release the frame before returning, so use a defer-style guard.
    let result = (|| -> Result<Vec<u8>, WindowsMcpError> {
        let desktop_resource = desktop_resource.ok_or_else(|| {
            WindowsMcpError::ScreenshotError(
                "AcquireNextFrame returned null desktop resource".into(),
            )
        })?;

        let gpu_texture: ID3D11Texture2D =
            desktop_resource.cast::<ID3D11Texture2D>().map_err(|e| {
                WindowsMcpError::ScreenshotError(format!(
                    "Desktop resource -> ID3D11Texture2D cast failed: {e}"
                ))
            })?;

        // Create a CPU-readable staging texture with the same dimensions.
        let staging_desc = D3D11_TEXTURE2D_DESC {
            Width: width,
            Height: height,
            MipLevels: 1,
            ArraySize: 1,
            Format: DXGI_FORMAT_B8G8R8A8_UNORM,
            SampleDesc: DXGI_SAMPLE_DESC {
                Count: 1,
                Quality: 0,
            },
            Usage: D3D11_USAGE_STAGING,
            BindFlags: D3D11_BIND_FLAG(0).0 as u32,
            CPUAccessFlags: D3D11_CPU_ACCESS_READ.0 as u32,
            MiscFlags: windows::Win32::Graphics::Direct3D11::D3D11_RESOURCE_MISC_FLAG(0).0 as u32,
        };

        let mut staging_texture: Option<ID3D11Texture2D> = None;
        unsafe {
            device
                .CreateTexture2D(&staging_desc, None, Some(&mut staging_texture))
                .map_err(|e| {
                    WindowsMcpError::ScreenshotError(format!(
                        "CreateTexture2D (staging) failed: {e}"
                    ))
                })?;
        }

        let staging_texture = staging_texture.ok_or_else(|| {
            WindowsMcpError::ScreenshotError(
                "CreateTexture2D returned null staging texture".into(),
            )
        })?;

        // Copy GPU texture -> staging texture.
        unsafe {
            context.CopyResource(&staging_texture, &gpu_texture);
        }

        // Map the staging texture to get a CPU pointer.
        let mut mapped = windows::Win32::Graphics::Direct3D11::D3D11_MAPPED_SUBRESOURCE::default();
        unsafe {
            context
                .Map(&staging_texture, 0, D3D11_MAP_READ, 0, Some(&mut mapped))
                .map_err(|e| {
                    WindowsMcpError::ScreenshotError(format!(
                        "ID3D11DeviceContext::Map failed: {e}"
                    ))
                })?;
        }

        // Copy pixels out of mapped memory row-by-row.
        // `mapped.RowPitch` may be larger than `width * 4` due to GPU
        // alignment padding; we must skip padding bytes on each row.
        let row_pitch = mapped.RowPitch as usize;
        let row_bytes = (width * 4) as usize;
        let mut pixels: Vec<u8> = Vec::with_capacity(row_bytes * height as usize);

        unsafe {
            let src_ptr = mapped.pData as *const u8;
            for row in 0..height as usize {
                let src_row = src_ptr.add(row * row_pitch);
                let src_slice = std::slice::from_raw_parts(src_row, row_bytes);
                pixels.extend_from_slice(src_slice);
            }
        }

        // Unmap before releasing the staging texture.
        unsafe {
            context.Unmap(&staging_texture, 0);
        }

        Ok(pixels)
    })();

    // Always release the acquired frame, even if pixel read failed.
    unsafe {
        let _ = duplication.ReleaseFrame();
    }

    result
}

// ---------------------------------------------------------------------------
// DXGI capture entry point
// ---------------------------------------------------------------------------

/// Capture the desktop for `monitor_index` via DXGI Output Duplication.
///
/// Returns raw BGRA pixel data.  This path requires a hardware or WARP
/// D3D11 device and fails inside pure Remote Desktop sessions without
/// GPU access.  Use [`capture_raw`] which automatically falls back to GDI.
fn capture_dxgi(monitor_index: u32) -> Result<ScreenshotData, WindowsMcpError> {
    let (device, context) = create_d3d11_device()?;

    let (adapter, output1, desc) = get_dxgi_output(monitor_index)?;

    // Derive monitor dimensions from the output descriptor.
    let desktop_rect: RECT = desc.DesktopCoordinates;
    let width = (desktop_rect.right - desktop_rect.left).unsigned_abs();
    let height = (desktop_rect.bottom - desktop_rect.top).unsigned_abs();

    if width == 0 || height == 0 {
        return Err(WindowsMcpError::ScreenshotError(format!(
            "Monitor {monitor_index} has zero-size desktop rect ({width}x{height})"
        )));
    }

    // DuplicateOutput requires the D3D11 device that was created against
    // the same adapter as the output.  We create the device fresh for
    // the correct adapter here.
    //
    // If the device was created against a different adapter (e.g. hardware
    // failed and we used WARP), create a new device for the correct adapter.
    let duplication: IDXGIOutputDuplication = {
        // Try to create device against the specific adapter owning this output.
        let mut specific_device: Option<ID3D11Device> = None;
        let mut specific_context: Option<ID3D11DeviceContext> = None;
        let mut fl = windows::Win32::Graphics::Direct3D::D3D_FEATURE_LEVEL_9_1;

        let hr = unsafe {
            D3D11CreateDevice(
                &adapter,
                windows::Win32::Graphics::Direct3D::D3D_DRIVER_TYPE_UNKNOWN,
                None,
                windows::Win32::Graphics::Direct3D11::D3D11_CREATE_DEVICE_FLAG(0),
                None,
                D3D11_SDK_VERSION,
                Some(&mut specific_device),
                Some(&mut fl),
                Some(&mut specific_context),
            )
        };

        // Use the adapter-specific device when possible; fall back to the
        // already-created WARP device otherwise.
        let (dup_device, _dup_context) = if hr.is_ok() {
            let d = specific_device.ok_or_else(|| {
                WindowsMcpError::ScreenshotError(
                    "D3D11CreateDevice (adapter) returned null device".into(),
                )
            })?;
            let c = specific_context.ok_or_else(|| {
                WindowsMcpError::ScreenshotError(
                    "D3D11CreateDevice (adapter) returned null context".into(),
                )
            })?;
            (d, c)
        } else {
            (device.clone(), context.clone())
        };

        unsafe {
            output1
                .DuplicateOutput(&dup_device)
                .map_err(|e| {
                    WindowsMcpError::ScreenshotError(format!("DuplicateOutput failed: {e}"))
                })?
        }
    };

    // Acquire and read one frame.
    let pixels = read_frame(&device, &context, &duplication, width, height)?;

    Ok(ScreenshotData {
        width,
        height,
        data: pixels,
    })
}

// ---------------------------------------------------------------------------
// GDI fallback capture
// ---------------------------------------------------------------------------

/// Capture the primary monitor using GDI `BitBlt`.
///
/// This fallback is used when DXGI Output Duplication is unavailable
/// (Remote Desktop sessions, some virtual machines, Windows Server
/// without a display driver).  It captures only the primary monitor
/// regardless of `monitor_index`.
///
/// Returns BGRA pixels (GDI DIBSection in `BI_RGB` 32-bit mode produces
/// `BGRA` layout with the alpha channel set to 0; we force alpha to 255).
fn capture_gdi(monitor_index: u32) -> Result<ScreenshotData, WindowsMcpError> {
    // GDI can only capture the primary monitor; warn if index > 0.
    if monitor_index > 0 {
        return Err(WindowsMcpError::ScreenshotError(format!(
            "GDI fallback does not support monitor index {monitor_index}; \
             only monitor 0 (primary) is supported"
        )));
    }

    let width = unsafe { GetSystemMetrics(SM_CXSCREEN) };
    let height = unsafe { GetSystemMetrics(SM_CYSCREEN) };

    if width <= 0 || height <= 0 {
        return Err(WindowsMcpError::ScreenshotError(format!(
            "GetSystemMetrics returned invalid screen size: {width}x{height}"
        )));
    }

    let (width, height) = (width as u32, height as u32);

    unsafe {
        // Get the screen DC.
        let screen_dc = GetDC(HWND(std::ptr::null_mut()));
        if screen_dc.is_invalid() {
            return Err(WindowsMcpError::ScreenshotError(
                "GetDC(NULL) failed".into(),
            ));
        }
        // RAII-style cleanup via a guard closure at the end.
        let result = (|| -> Result<ScreenshotData, WindowsMcpError> {
            let mem_dc = CreateCompatibleDC(screen_dc);
            if mem_dc.is_invalid() {
                return Err(WindowsMcpError::ScreenshotError(
                    "CreateCompatibleDC failed".into(),
                ));
            }
            let bitmap = CreateCompatibleBitmap(screen_dc, width as i32, height as i32);
            if bitmap.is_invalid() {
                let _ = DeleteDC(mem_dc);
                return Err(WindowsMcpError::ScreenshotError(
                    "CreateCompatibleBitmap failed".into(),
                ));
            }

            let old_bitmap = SelectObject(mem_dc, bitmap);

            // Copy from screen DC to memory DC.
            BitBlt(mem_dc, 0, 0, width as i32, height as i32, screen_dc, 0, 0, SRCCOPY)
                .map_err(|e| {
                    SelectObject(mem_dc, old_bitmap);
                    let _ = DeleteObject(bitmap);
                    let _ = DeleteDC(mem_dc);
                    WindowsMcpError::ScreenshotError(format!("BitBlt failed: {e}"))
                })?;

            // Retrieve pixels in 32-bit BGRA format.
            let pixel_count = (width * height) as usize;
            let mut pixels = vec![0u8; pixel_count * 4];

            let bmi = BITMAPINFO {
                bmiHeader: BITMAPINFOHEADER {
                    biSize: std::mem::size_of::<BITMAPINFOHEADER>() as u32,
                    biWidth: width as i32,
                    // Negative height = top-down bitmap (row 0 at top).
                    biHeight: -(height as i32),
                    biPlanes: 1,
                    biBitCount: 32,
                    biCompression: BI_RGB.0,
                    biSizeImage: 0,
                    biXPelsPerMeter: 0,
                    biYPelsPerMeter: 0,
                    biClrUsed: 0,
                    biClrImportant: 0,
                },
                bmiColors: [Default::default()],
            };

            let lines = GetDIBits(
                mem_dc,
                bitmap,
                0,
                height,
                Some(pixels.as_mut_ptr() as *mut _),
                &bmi as *const _ as *mut _,
                DIB_RGB_COLORS,
            );

            SelectObject(mem_dc, old_bitmap);
            let _ = DeleteObject(bitmap);
            let _ = DeleteDC(mem_dc);

            if lines == 0 {
                return Err(WindowsMcpError::ScreenshotError("GetDIBits failed".into()));
            }

            // GDI BI_RGB 32-bit has alpha = 0; set it to 255 (fully opaque).
            for chunk in pixels.chunks_exact_mut(4) {
                chunk[3] = 255;
            }

            Ok(ScreenshotData {
                width,
                height,
                data: pixels,
            })
        })();

        ReleaseDC(HWND(std::ptr::null_mut()), screen_dc);
        result
    }
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/// Capture the desktop for `monitor_index` and return raw BGRA pixel data.
///
/// Attempts DXGI Output Duplication first for high-performance headless
/// capture.  Falls back to GDI `BitBlt` automatically if DXGI is
/// unavailable (Remote Desktop, VM without GPU, etc.).
///
/// # Parameters
///
/// - `monitor_index`: Zero-based index of the monitor to capture.
///   Pass `0` for the primary monitor.
///
/// # Returns
///
/// A [`ScreenshotData`] with `width`, `height`, and `data` (BGRA bytes).
///
/// # Errors
///
/// Returns [`crate::errors::WindowsMcpError::ScreenshotError`] if both
/// DXGI and GDI capture fail.
///
/// # Examples
///
/// ```no_run
/// use wmcp_core::screenshot::capture_raw;
///
/// let frame = capture_raw(0).expect("capture failed");
/// assert_eq!(frame.data.len(), (frame.width * frame.height * 4) as usize);
/// ```
pub fn capture_raw(monitor_index: u32) -> Result<ScreenshotData, WindowsMcpError> {
    match capture_dxgi(monitor_index) {
        Ok(data) => Ok(data),
        Err(dxgi_err) => {
            log::warn!(
                "DXGI capture failed for monitor {monitor_index} ({dxgi_err}); \
                 falling back to GDI BitBlt"
            );
            capture_gdi(monitor_index)
        }
    }
}

/// Capture the desktop for `monitor_index` and encode it as a PNG.
///
/// Internally calls [`capture_raw`] and encodes the BGRA pixel data
/// using the [`image`] crate.  The PNG is returned as a `Vec<u8>` in
/// memory -- write it to a file or send it over the network directly.
///
/// # Parameters
///
/// - `monitor_index`: Zero-based index of the monitor to capture.
///
/// # Returns
///
/// A `Vec<u8>` containing a valid PNG file.
///
/// # Errors
///
/// Returns [`crate::errors::WindowsMcpError::ScreenshotError`] if capture
/// or PNG encoding fails.
///
/// # Examples
///
/// ```no_run
/// use wmcp_core::screenshot::capture_png;
///
/// let png = capture_png(0).expect("PNG capture failed");
/// std::fs::write("desktop.png", &png).unwrap();
/// ```
pub fn capture_png(monitor_index: u32) -> Result<Vec<u8>, WindowsMcpError> {
    let frame = capture_raw(monitor_index)?;

    // Convert BGRA -> RGBA for the `image` crate (which uses RGBA layout).
    let rgba_pixels: Vec<u8> = frame
        .data
        .chunks_exact(4)
        .flat_map(|px| {
            // px = [B, G, R, A]
            [px[2], px[1], px[0], px[3]]
        })
        .collect();

    let img = image::RgbaImage::from_raw(frame.width, frame.height, rgba_pixels)
        .ok_or_else(|| {
            WindowsMcpError::ScreenshotError(
                "image::RgbaImage::from_raw failed: buffer size mismatch".into(),
            )
        })?;

    let mut buf: Vec<u8> = Vec::new();
    let mut cursor = std::io::Cursor::new(&mut buf);

    img.write_to(&mut cursor, image::ImageFormat::Png)
        .map_err(|e| {
            WindowsMcpError::ScreenshotError(format!("PNG encoding failed: {e}"))
        })?;

    Ok(buf)
}
