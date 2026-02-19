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
//! 4. `AcquireNextFrame` returns an `IDXGIResource` backed by a GPU texture.
//! 5. Create a CPU-readable staging texture (`D3D11_USAGE_STAGING`), copy
//!    the desktop frame into it, then map it with `D3D11_MAP_READ` to
//!    obtain a `*const u8` pointer to BGRA pixels.
//! 6. Release the frame and duplicate interface after reading.
//!
//! # Thread safety
//!
//! All DXGI / D3D11 interfaces are COM objects.  This module creates them
//! fresh on every call -- there is no shared global state.  Each call must
//! be made from a thread with a valid COM apartment (call
//! [`crate::com::COMGuard::init`] before invoking these functions from a
//! new thread).
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
    D3D11_CPU_ACCESS_READ, D3D11_MAP_READ, D3D11_MAPPED_SUBRESOURCE, D3D11_SDK_VERSION,
    D3D11_TEXTURE2D_DESC, D3D11_USAGE_STAGING,
};
use windows::Win32::Graphics::Dxgi::Common::{DXGI_FORMAT_B8G8R8A8_UNORM, DXGI_SAMPLE_DESC};
use windows::Win32::Graphics::Dxgi::{
    CreateDXGIFactory1, IDXGIAdapter, IDXGIFactory1, IDXGIOutput, IDXGIOutput1,
    IDXGIOutputDuplication, IDXGIResource, DXGI_OUTDUPL_FRAME_INFO,
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
    let mut feature_level = windows::Win32::Graphics::Direct3D::D3D_FEATURE_LEVEL_9_1;

    // Try hardware adapter first.
    let hw_result = unsafe {
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

    if hw_result.is_ok() {
        let dev = device.ok_or_else(|| {
            WindowsMcpError::ScreenshotError(
                "D3D11CreateDevice (HW) succeeded but returned null device".into(),
            )
        })?;
        let ctx = context.ok_or_else(|| {
            WindowsMcpError::ScreenshotError(
                "D3D11CreateDevice (HW) succeeded but returned null context".into(),
            )
        })?;
        return Ok((dev, ctx));
    }

    // Hardware failed -- try WARP software renderer.
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

    let dev = device2.ok_or_else(|| {
        WindowsMcpError::ScreenshotError(
            "D3D11CreateDevice (WARP) returned null device".into(),
        )
    })?;
    let ctx = context2.ok_or_else(|| {
        WindowsMcpError::ScreenshotError(
            "D3D11CreateDevice (WARP) returned null context".into(),
        )
    })?;
    Ok((dev, ctx))
}

/// Enumerate DXGI outputs (monitors) and return the `IDXGIOutput1` for
/// `monitor_index`, plus the adapter that owns it and the monitor
/// desktop coordinates.
fn get_dxgi_output(
    monitor_index: u32,
) -> Result<(IDXGIAdapter, IDXGIOutput1, RECT), WindowsMcpError> {
    let factory: IDXGIFactory1 = unsafe {
        CreateDXGIFactory1().map_err(|e| {
            WindowsMcpError::ScreenshotError(format!("CreateDXGIFactory1 failed: {e}"))
        })?
    };

    let mut global_output_index: u32 = 0;

    // Walk adapters (graphics cards) in enumeration order.
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

                // GetDesc returns Result<DXGI_OUTPUT_DESC> in windows 0.58.
                let desc = unsafe {
                    output1.GetDesc().map_err(|e| {
                        WindowsMcpError::ScreenshotError(format!(
                            "IDXGIOutput1::GetDesc failed: {e}"
                        ))
                    })?
                };

                return Ok((adapter, output1, desc.DesktopCoordinates));
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

/// Acquire one frame from `duplication`, copy it into a CPU-readable
/// staging texture, and return the raw BGRA pixel bytes.
///
/// The device/context pair must have been created against the same DXGI
/// adapter as the output used to create `duplication`.
fn read_frame(
    device: &ID3D11Device,
    context: &ID3D11DeviceContext,
    duplication: &IDXGIOutputDuplication,
    width: u32,
    height: u32,
) -> Result<Vec<u8>, WindowsMcpError> {
    // AcquireNextFrame blocks until a new frame is available.
    // 500ms timeout is ample for a 60Hz display (~16ms between frames).
    let mut frame_info = DXGI_OUTDUPL_FRAME_INFO::default();
    // AcquireNextFrame takes *mut Option<IDXGIResource> -- must use a raw ptr.
    let mut desktop_resource: Option<IDXGIResource> = None;

    unsafe {
        duplication
            .AcquireNextFrame(
                500,
                std::ptr::addr_of_mut!(frame_info),
                std::ptr::addr_of_mut!(desktop_resource),
            )
            .map_err(|e| {
                WindowsMcpError::ScreenshotError(format!("AcquireNextFrame failed: {e}"))
            })?;
    }

    // We must call ReleaseFrame before returning -- even on error paths.
    // Implement with a defer-style closure.
    let pixel_result = (|| -> Result<Vec<u8>, WindowsMcpError> {
        let resource = desktop_resource.ok_or_else(|| {
            WindowsMcpError::ScreenshotError(
                "AcquireNextFrame returned a null desktop resource".into(),
            )
        })?;

        // QueryInterface the IDXGIResource to get an ID3D11Texture2D.
        let gpu_texture: ID3D11Texture2D = resource.cast::<ID3D11Texture2D>().map_err(|e| {
            WindowsMcpError::ScreenshotError(format!(
                "IDXGIResource -> ID3D11Texture2D cast failed: {e}"
            ))
        })?;

        // Create a CPU-readable staging texture with matching dimensions.
        // BindFlags / CPUAccessFlags / MiscFlags are plain u32 in this struct.
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
            BindFlags: 0,
            CPUAccessFlags: D3D11_CPU_ACCESS_READ.0 as u32,
            MiscFlags: 0,
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
                "CreateTexture2D returned a null staging texture".into(),
            )
        })?;

        // Copy the GPU-resident desktop texture into the CPU-readable staging
        // texture.  This is a GPU-to-GPU blit; the CPU sees the result after
        // Map().
        unsafe {
            context.CopyResource(&staging_texture, &gpu_texture);
        }

        // Map the staging texture to obtain a CPU pointer to the pixel data.
        let mut mapped = D3D11_MAPPED_SUBRESOURCE::default();
        unsafe {
            context
                .Map(
                    &staging_texture,
                    0,
                    D3D11_MAP_READ,
                    0,
                    Some(&mut mapped),
                )
                .map_err(|e| {
                    WindowsMcpError::ScreenshotError(format!(
                        "ID3D11DeviceContext::Map failed: {e}"
                    ))
                })?;
        }

        // Copy pixels out row-by-row.
        // `mapped.RowPitch` >= `width * 4` due to GPU alignment padding;
        // skip the padding bytes at the end of each row.
        let row_pitch = mapped.RowPitch as usize;
        let row_bytes = (width as usize) * 4;
        let mut pixels: Vec<u8> = Vec::with_capacity(row_bytes * height as usize);

        unsafe {
            let src_ptr = mapped.pData as *const u8;
            for row in 0..height as usize {
                let row_start = src_ptr.add(row * row_pitch);
                let src_slice = std::slice::from_raw_parts(row_start, row_bytes);
                pixels.extend_from_slice(src_slice);
            }
        }

        // Unmap before the staging texture is dropped.
        unsafe {
            context.Unmap(&staging_texture, 0);
        }

        Ok(pixels)
    })();

    // Always release the acquired DXGI frame.
    unsafe {
        let _ = duplication.ReleaseFrame();
    }

    pixel_result
}

// ---------------------------------------------------------------------------
// DXGI capture entry point
// ---------------------------------------------------------------------------

/// Capture the desktop for `monitor_index` via DXGI Output Duplication.
///
/// Returns raw BGRA pixel data or a [`WindowsMcpError::ScreenshotError`].
/// This path requires a D3D11 device and fails in pure Remote Desktop
/// sessions without GPU passthrough.  The public [`capture_raw`] falls
/// back to GDI automatically.
fn capture_dxgi(monitor_index: u32) -> Result<ScreenshotData, WindowsMcpError> {
    // Retrieve the target adapter/output so we can bind DuplicateOutput to
    // the correct device.
    let (adapter, output1, desktop_rect) = get_dxgi_output(monitor_index)?;

    let width = (desktop_rect.right - desktop_rect.left).unsigned_abs();
    let height = (desktop_rect.bottom - desktop_rect.top).unsigned_abs();

    if width == 0 || height == 0 {
        return Err(WindowsMcpError::ScreenshotError(format!(
            "Monitor {monitor_index} has zero-size desktop rect ({width}x{height})"
        )));
    }

    // Create the D3D11 device against the specific adapter that owns the
    // output.  DuplicateOutput requires the device and output to share
    // the same DXGI adapter; passing D3D_DRIVER_TYPE_UNKNOWN with an
    // explicit adapter achieves this.
    let mut device_opt: Option<ID3D11Device> = None;
    let mut context_opt: Option<ID3D11DeviceContext> = None;
    let mut feature_level = windows::Win32::Graphics::Direct3D::D3D_FEATURE_LEVEL_9_1;

    let hr = unsafe {
        D3D11CreateDevice(
            &adapter,
            windows::Win32::Graphics::Direct3D::D3D_DRIVER_TYPE_UNKNOWN,
            None,
            windows::Win32::Graphics::Direct3D11::D3D11_CREATE_DEVICE_FLAG(0),
            None,
            D3D11_SDK_VERSION,
            Some(&mut device_opt),
            Some(&mut feature_level),
            Some(&mut context_opt),
        )
    };

    // Fall back to a generic hardware/WARP device if adapter-specific
    // creation fails (can happen on some hybrid GPU configurations).
    let (device, context) = if hr.is_ok() {
        (
            device_opt.ok_or_else(|| {
                WindowsMcpError::ScreenshotError(
                    "D3D11CreateDevice (adapter) returned null device".into(),
                )
            })?,
            context_opt.ok_or_else(|| {
                WindowsMcpError::ScreenshotError(
                    "D3D11CreateDevice (adapter) returned null context".into(),
                )
            })?,
        )
    } else {
        create_d3d11_device()?
    };

    // Open the output duplication session.
    let duplication: IDXGIOutputDuplication = unsafe {
        output1
            .DuplicateOutput(&device)
            .map_err(|e| {
                WindowsMcpError::ScreenshotError(format!("DuplicateOutput failed: {e}"))
            })?
    };

    // Capture one frame.
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
/// Used when DXGI Output Duplication is unavailable (Remote Desktop
/// sessions, virtual machines without GPU access, Windows Server SKUs
/// that lack a hardware display driver).  Only the primary monitor
/// (`monitor_index == 0`) is supported.
///
/// GDI `BI_RGB` 32-bit mode stores pixels as BGRA with alpha == 0;
/// this function sets alpha to 255 (fully opaque) before returning.
fn capture_gdi(monitor_index: u32) -> Result<ScreenshotData, WindowsMcpError> {
    if monitor_index > 0 {
        return Err(WindowsMcpError::ScreenshotError(format!(
            "GDI fallback does not support monitor index {monitor_index}; \
             only monitor 0 (primary) is supported via GDI BitBlt"
        )));
    }

    let width_i = unsafe { GetSystemMetrics(SM_CXSCREEN) };
    let height_i = unsafe { GetSystemMetrics(SM_CYSCREEN) };

    if width_i <= 0 || height_i <= 0 {
        return Err(WindowsMcpError::ScreenshotError(format!(
            "GetSystemMetrics returned invalid screen size: {width_i}x{height_i}"
        )));
    }

    let width = width_i as u32;
    let height = height_i as u32;

    unsafe {
        let screen_dc = GetDC(HWND(std::ptr::null_mut()));
        if screen_dc.is_invalid() {
            return Err(WindowsMcpError::ScreenshotError(
                "GetDC(NULL) returned an invalid DC".into(),
            ));
        }

        // Use a nested closure so we can call ReleaseDC unconditionally.
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

            let bitblt_result =
                BitBlt(mem_dc, 0, 0, width as i32, height as i32, screen_dc, 0, 0, SRCCOPY);

            if bitblt_result.is_err() {
                SelectObject(mem_dc, old_bitmap);
                let _ = DeleteObject(bitmap);
                let _ = DeleteDC(mem_dc);
                return Err(WindowsMcpError::ScreenshotError("BitBlt failed".into()));
            }

            // GetDIBits expects *mut BITMAPINFO.
            let mut pixels = vec![0u8; (width * height * 4) as usize];
            let mut bmi = BITMAPINFO {
                bmiHeader: BITMAPINFOHEADER {
                    biSize: std::mem::size_of::<BITMAPINFOHEADER>() as u32,
                    biWidth: width as i32,
                    // Negative height = top-down row order (row 0 at top).
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
                Some(pixels.as_mut_ptr().cast()),
                &mut bmi,
                DIB_RGB_COLORS,
            );

            SelectObject(mem_dc, old_bitmap);
            let _ = DeleteObject(bitmap);
            let _ = DeleteDC(mem_dc);

            if lines == 0 {
                return Err(WindowsMcpError::ScreenshotError("GetDIBits returned 0".into()));
            }

            // GDI BI_RGB 32-bit sets alpha to 0; force it to 255 (opaque).
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
/// `data.len() == width * height * 4` is always true on success.
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
/// memory -- write it to a file or transmit it directly.
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

    // Convert BGRA -> RGBA for the `image` crate (its RgbaImage uses RGBA).
    let rgba_pixels: Vec<u8> = frame
        .data
        .chunks_exact(4)
        .flat_map(|px| {
            // px layout: [B, G, R, A]
            [px[2], px[1], px[0], px[3]]
        })
        .collect();

    let img = image::RgbaImage::from_raw(frame.width, frame.height, rgba_pixels)
        .ok_or_else(|| {
            WindowsMcpError::ScreenshotError(
                "image::RgbaImage::from_raw failed: pixel buffer size mismatch".into(),
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
