use anyhow::{Context, Result, anyhow};
use image::{ImageBuffer, Rgba};
use rmcp::handler::server::router::tool::ToolRouter;
use rmcp::handler::server::wrapper::Parameters;
use rmcp::model::{CallToolResult, Content, ErrorCode, ServerCapabilities, ServerInfo};
use rmcp::schemars::JsonSchema;
use rmcp::transport::stdio;
use rmcp::{ErrorData as McpError, ServerHandler, ServiceExt, tool, tool_handler, tool_router};
use serde::{Deserialize, Serialize};
use std::ffi::OsString;
use std::ffi::c_void;
use std::os::windows::ffi::OsStringExt;
use std::path::{Path, PathBuf};
use uuid::Uuid;
use windows::Win32::Foundation::{BOOL, CloseHandle, HWND, LPARAM, RECT};
use windows::Win32::Graphics::Gdi::{
    BI_RGB, BITMAPINFO, BITMAPINFOHEADER, BitBlt, CreateCompatibleBitmap, CreateCompatibleDC,
    DIB_RGB_COLORS, DeleteDC, DeleteObject, GetDC, GetDIBits, HGDIOBJ, ReleaseDC, SRCCOPY,
    SelectObject,
};
use windows::Win32::System::Threading::{
    OpenProcess, PROCESS_NAME_FORMAT, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_VM_READ,
    QueryFullProcessImageNameW,
};
use windows::Win32::UI::WindowsAndMessaging::{
    EnumWindows, GetWindowRect, GetWindowTextLengthW, GetWindowTextW, GetWindowThreadProcessId,
    IsWindowVisible,
};
use windows::core::PWSTR;

#[derive(Debug, Clone, Serialize)]
struct WindowInfo {
    hwnd: i64,
    process_name: String,
    title: String,
    pid: u32,
}

#[derive(Debug, Clone, Copy)]
struct RelativeRect {
    x: u32,
    y: u32,
    width: u32,
    height: u32,
}

#[derive(Debug, Clone)]
struct CaptureServer {
    tool_router: ToolRouter<Self>,
}

#[derive(Debug, Deserialize, JsonSchema)]
#[serde(untagged)]
enum HwndInput {
    Number(i64),
    String(String),
}

#[derive(Debug, Deserialize, JsonSchema)]
struct CaptureByProcessNameArgs {
    process_name: String,
    save_path: String,
    x: Option<u32>,
    y: Option<u32>,
    width: Option<u32>,
    height: Option<u32>,
}

#[derive(Debug, Deserialize, JsonSchema)]
struct CaptureByHwndArgs {
    #[schemars(description = "Window handle, decimal integer or hex string like 0x1A2B")]
    hwnd: HwndInput,
    save_path: String,
    x: Option<u32>,
    y: Option<u32>,
    width: Option<u32>,
    height: Option<u32>,
}

#[tool_router]
impl CaptureServer {
    fn new() -> Self {
        Self {
            tool_router: Self::tool_router(),
        }
    }

    #[tool(description = "Get visible top-level windows and return hwnd + process name + title")]
    fn list_process_windows(&self) -> Result<String, McpError> {
        let windows = list_process_windows().map_err(to_mcp_error)?;
        serde_json::to_string_pretty(&windows).map_err(to_mcp_error)
    }

    #[tool(description = "Capture screenshot by process name and save to path")]
    fn capture_by_process_name(
        &self,
        Parameters(args): Parameters<CaptureByProcessNameArgs>,
    ) -> Result<CallToolResult, McpError> {
        let rect = relative_rect_from_parts(args.x, args.y, args.width, args.height);

        let windows = list_process_windows().map_err(to_mcp_error)?;
        let matched: Vec<WindowInfo> = windows
            .into_iter()
            .filter(|w| w.process_name.eq_ignore_ascii_case(&args.process_name))
            .collect();
        if matched.is_empty() {
            return Err(to_mcp_error(format!(
                "no visible window found for process_name={}",
                args.process_name
            )));
        }

        let target = matched
            .iter()
            .find(|w| !w.title.trim().is_empty())
            .cloned()
            .unwrap_or_else(|| matched[0].clone());

        let saved = capture_window_to_path(hwnd_from_i64(target.hwnd), &args.save_path, rect)
            .map_err(to_mcp_error)?;

        Ok(CallToolResult::success(vec![Content::text(format!(
            "saved: {}",
            saved.display()
        ))]))
    }

    #[tool(description = "Capture screenshot by window handle and save to path")]
    fn capture_by_hwnd(
        &self,
        Parameters(args): Parameters<CaptureByHwndArgs>,
    ) -> Result<CallToolResult, McpError> {
        let hwnd = parse_hwnd(&args.hwnd).map_err(to_mcp_error)?;
        let rect = relative_rect_from_parts(args.x, args.y, args.width, args.height);
        let saved = capture_window_to_path(hwnd, &args.save_path, rect).map_err(to_mcp_error)?;

        Ok(CallToolResult::success(vec![Content::text(format!(
            "saved: {}",
            saved.display()
        ))]))
    }
}

#[tool_handler]
impl ServerHandler for CaptureServer {
    fn get_info(&self) -> ServerInfo {
        ServerInfo::new(ServerCapabilities::builder().enable_tools().build())
            .with_instructions("Capture windows on Windows by process name or hwnd".to_string())
    }
}

#[tokio::main]
async fn main() -> Result<()> {
    let service = CaptureServer::new().serve(stdio()).await?;
    service.waiting().await?;
    Ok(())
}

fn to_mcp_error<E: std::fmt::Display>(error: E) -> McpError {
    McpError::new(ErrorCode::INTERNAL_ERROR, error.to_string(), None)
}

fn hwnd_from_i64(value: i64) -> HWND {
    HWND(value as usize as *mut c_void)
}

fn parse_hwnd(value: &HwndInput) -> Result<HWND> {
    match value {
        HwndInput::Number(v) => Ok(hwnd_from_i64(*v)),
        HwndInput::String(s) => {
            let trimmed = s.trim();
            let parsed = if let Some(hex) = trimmed
                .strip_prefix("0x")
                .or_else(|| trimmed.strip_prefix("0X"))
            {
                i64::from_str_radix(hex, 16)
                    .with_context(|| format!("invalid hex hwnd string: {trimmed}"))?
            } else {
                trimmed
                    .parse::<i64>()
                    .with_context(|| format!("invalid hwnd string: {trimmed}"))?
            };
            Ok(hwnd_from_i64(parsed))
        }
    }
}

fn relative_rect_from_parts(
    x: Option<u32>,
    y: Option<u32>,
    width: Option<u32>,
    height: Option<u32>,
) -> Option<RelativeRect> {
    if x.is_none() && y.is_none() && width.is_none() && height.is_none() {
        return None;
    }

    Some(RelativeRect {
        x: x.unwrap_or(0),
        y: y.unwrap_or(0),
        width: width.unwrap_or(0),
        height: height.unwrap_or(0),
    })
}

fn capture_window_to_path(
    hwnd: HWND,
    save_path: &str,
    requested: Option<RelativeRect>,
) -> Result<PathBuf> {
    let (img_w, img_h, rgba) = capture_window_rgba(hwnd)?;
    let rect = normalize_relative_rect(requested, img_w, img_h)?;
    let cropped = crop_rgba(&rgba, img_w, img_h, rect);

    let output = resolve_output_path(save_path)?;
    if let Some(parent) = output.parent() {
        std::fs::create_dir_all(parent)
            .with_context(|| format!("failed to create output directory: {}", parent.display()))?;
    }

    let img = ImageBuffer::<Rgba<u8>, Vec<u8>>::from_raw(rect.width, rect.height, cropped)
        .ok_or_else(|| anyhow!("failed to assemble image buffer"))?;
    img.save(&output)
        .with_context(|| format!("failed to save png: {}", output.display()))?;

    Ok(output)
}

fn normalize_relative_rect(
    requested: Option<RelativeRect>,
    img_w: u32,
    img_h: u32,
) -> Result<RelativeRect> {
    match requested {
        None => Ok(RelativeRect {
            x: 0,
            y: 0,
            width: img_w,
            height: img_h,
        }),
        Some(r) => {
            if r.x >= img_w || r.y >= img_h {
                return Err(anyhow!(
                    "x/y out of bounds: x={}, y={}, image={}x{}",
                    r.x,
                    r.y,
                    img_w,
                    img_h
                ));
            }
            let width = if r.width == 0 { img_w - r.x } else { r.width };
            let height = if r.height == 0 { img_h - r.y } else { r.height };

            if r.x.saturating_add(width) > img_w || r.y.saturating_add(height) > img_h {
                return Err(anyhow!(
                    "requested crop out of bounds: x={}, y={}, w={}, h={}, image={}x{}",
                    r.x,
                    r.y,
                    width,
                    height,
                    img_w,
                    img_h
                ));
            }

            Ok(RelativeRect {
                x: r.x,
                y: r.y,
                width,
                height,
            })
        }
    }
}

fn crop_rgba(source: &[u8], src_w: u32, _src_h: u32, rect: RelativeRect) -> Vec<u8> {
    let src_stride = src_w as usize * 4;
    let mut out = vec![0u8; rect.width as usize * rect.height as usize * 4];

    for row in 0..rect.height as usize {
        let src_row = rect.y as usize + row;
        let src_start = src_row * src_stride + rect.x as usize * 4;
        let src_end = src_start + rect.width as usize * 4;
        let dst_start = row * rect.width as usize * 4;
        let dst_end = dst_start + rect.width as usize * 4;
        out[dst_start..dst_end].copy_from_slice(&source[src_start..src_end]);
    }

    out
}

fn resolve_output_path(input: &str) -> Result<PathBuf> {
    let p = PathBuf::from(input);

    let file_path = if p.exists() && p.is_dir() {
        p.join(format!("{}.png", Uuid::new_v4()))
    } else if !p.exists() && p.extension().is_none() {
        std::fs::create_dir_all(&p)
            .with_context(|| format!("failed to create directory path: {}", p.display()))?;
        p.join(format!("{}.png", Uuid::new_v4()))
    } else {
        p
    };

    Ok(file_path)
}

fn list_process_windows() -> Result<Vec<WindowInfo>> {
    let mut hwnds: Vec<HWND> = Vec::new();

    unsafe extern "system" fn enum_windows_proc(hwnd: HWND, lparam: LPARAM) -> BOOL {
        let vec_ptr = lparam.0 as *mut Vec<HWND>;
        if vec_ptr.is_null() {
            return BOOL(0);
        }
        if unsafe { IsWindowVisible(hwnd) }.as_bool() {
            // Safety: vec_ptr comes from a valid mutable reference for the duration of EnumWindows.
            unsafe {
                (*vec_ptr).push(hwnd);
            }
        }
        BOOL(1)
    }

    unsafe {
        EnumWindows(
            Some(enum_windows_proc),
            LPARAM((&mut hwnds as *mut Vec<HWND>) as isize),
        )
        .context("EnumWindows failed")?;
    }

    let mut out = Vec::new();
    for hwnd in hwnds {
        let title = get_window_title(hwnd)?;
        if title.trim().is_empty() {
            continue;
        }

        let pid = get_window_pid(hwnd);
        let process_name = get_process_name_by_pid(pid).unwrap_or_else(|_| format!("pid-{pid}"));
        out.push(WindowInfo {
            hwnd: hwnd.0 as usize as i64,
            process_name,
            title,
            pid,
        });
    }

    Ok(out)
}

fn get_window_pid(hwnd: HWND) -> u32 {
    let mut pid = 0u32;
    unsafe {
        GetWindowThreadProcessId(hwnd, Some(&mut pid));
    }
    pid
}

fn get_window_title(hwnd: HWND) -> Result<String> {
    let len = unsafe { GetWindowTextLengthW(hwnd) };
    if len <= 0 {
        return Ok(String::new());
    }

    let mut buffer = vec![0u16; (len + 1) as usize];
    let copied = unsafe { GetWindowTextW(hwnd, &mut buffer) };
    if copied <= 0 {
        return Ok(String::new());
    }

    let os = OsString::from_wide(&buffer[..copied as usize]);
    Ok(os.to_string_lossy().to_string())
}

fn get_process_name_by_pid(pid: u32) -> Result<String> {
    let handle = unsafe {
        OpenProcess(
            PROCESS_QUERY_LIMITED_INFORMATION | PROCESS_VM_READ,
            false,
            pid,
        )
    }
    .with_context(|| format!("OpenProcess failed for pid={pid}"))?;

    let mut size: u32 = 1024;
    let mut buffer = vec![0u16; size as usize];
    let ok = unsafe {
        QueryFullProcessImageNameW(
            handle,
            PROCESS_NAME_FORMAT(0),
            PWSTR(buffer.as_mut_ptr()),
            &mut size,
        )
    };
    unsafe {
        let _ = CloseHandle(handle);
    }

    if ok.is_err() {
        return Err(anyhow!("QueryFullProcessImageNameW failed for pid={pid}"));
    }

    let full = OsString::from_wide(&buffer[..size as usize])
        .to_string_lossy()
        .to_string();
    let name = Path::new(&full)
        .file_name()
        .map(|s| s.to_string_lossy().to_string())
        .unwrap_or(full);
    Ok(name)
}

fn capture_window_rgba(hwnd: HWND) -> Result<(u32, u32, Vec<u8>)> {
    let mut rect = RECT::default();
    unsafe { GetWindowRect(hwnd, &mut rect).context("GetWindowRect failed")? };

    let width = (rect.right - rect.left).max(0) as u32;
    let height = (rect.bottom - rect.top).max(0) as u32;
    if width == 0 || height == 0 {
        return Err(anyhow!("window has zero size"));
    }

    let hdc_screen = unsafe { GetDC(HWND(std::ptr::null_mut())) };
    if hdc_screen.0.is_null() {
        return Err(anyhow!("GetDC failed"));
    }

    let hdc_mem = unsafe { CreateCompatibleDC(hdc_screen) };
    if hdc_mem.0.is_null() {
        unsafe {
            let _ = ReleaseDC(HWND(std::ptr::null_mut()), hdc_screen);
        }
        return Err(anyhow!("CreateCompatibleDC failed"));
    }

    let hbitmap = unsafe { CreateCompatibleBitmap(hdc_screen, width as i32, height as i32) };
    if hbitmap.0.is_null() {
        unsafe {
            let _ = DeleteDC(hdc_mem);
            let _ = ReleaseDC(HWND(std::ptr::null_mut()), hdc_screen);
        }
        return Err(anyhow!("CreateCompatibleBitmap failed"));
    }

    let old_obj = unsafe { SelectObject(hdc_mem, HGDIOBJ(hbitmap.0)) };

    let bitblt_ok = unsafe {
        BitBlt(
            hdc_mem,
            0,
            0,
            width as i32,
            height as i32,
            hdc_screen,
            rect.left,
            rect.top,
            SRCCOPY,
        )
    }
    .is_ok();
    if !bitblt_ok {
        unsafe {
            let _ = SelectObject(hdc_mem, old_obj);
            let _ = DeleteObject(HGDIOBJ(hbitmap.0));
            let _ = DeleteDC(hdc_mem);
            let _ = ReleaseDC(HWND(std::ptr::null_mut()), hdc_screen);
        }
        return Err(anyhow!("BitBlt failed"));
    }

    let mut bi = BITMAPINFO {
        bmiHeader: BITMAPINFOHEADER {
            biSize: std::mem::size_of::<BITMAPINFOHEADER>() as u32,
            biWidth: width as i32,
            biHeight: -(height as i32),
            biPlanes: 1,
            biBitCount: 32,
            biCompression: BI_RGB.0,
            ..Default::default()
        },
        ..Default::default()
    };

    let mut raw = vec![0u8; width as usize * height as usize * 4];
    let rows = unsafe {
        GetDIBits(
            hdc_mem,
            hbitmap,
            0,
            height,
            Some(raw.as_mut_ptr() as *mut _),
            &mut bi,
            DIB_RGB_COLORS,
        )
    };

    unsafe {
        let _ = SelectObject(hdc_mem, old_obj);
        let _ = DeleteObject(HGDIOBJ(hbitmap.0));
        let _ = DeleteDC(hdc_mem);
        let _ = ReleaseDC(HWND(std::ptr::null_mut()), hdc_screen);
    }

    if rows == 0 {
        return Err(anyhow!("GetDIBits failed"));
    }

    for px in raw.chunks_exact_mut(4) {
        px.swap(0, 2);
    }

    Ok((width, height, raw))
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::fs;

    #[test]
    fn parse_hwnd_accepts_number_and_hex_string() {
        let from_number = parse_hwnd(&HwndInput::Number(4660)).expect("number hwnd should parse");
        let from_hex =
            parse_hwnd(&HwndInput::String("0x1234".to_string())).expect("hex hwnd should parse");
        let from_decimal = parse_hwnd(&HwndInput::String("4660".to_string()))
            .expect("decimal string hwnd should parse");

        assert_eq!(from_number.0 as usize, 0x1234);
        assert_eq!(from_hex.0 as usize, 0x1234);
        assert_eq!(from_decimal.0 as usize, 0x1234);
    }

    #[test]
    fn parse_hwnd_invalid_string_fails() {
        let err = parse_hwnd(&HwndInput::String("not-a-handle".to_string()))
            .expect_err("invalid hwnd string should fail");
        assert!(err.to_string().contains("invalid hwnd string"));
    }

    #[test]
    fn relative_rect_from_parts_none_when_all_missing() {
        let rect = relative_rect_from_parts(None, None, None, None);
        assert!(rect.is_none());
    }

    #[test]
    fn relative_rect_from_parts_defaults_missing_values() {
        let rect =
            relative_rect_from_parts(Some(5), Some(7), None, None).expect("rect should be present");
        assert_eq!(rect.x, 5);
        assert_eq!(rect.y, 7);
        assert_eq!(rect.width, 0);
        assert_eq!(rect.height, 0);
    }

    #[test]
    fn normalize_relative_rect_defaults_to_full_image() {
        let rect = normalize_relative_rect(None, 640, 480).expect("default rect should be valid");
        assert_eq!(rect.x, 0);
        assert_eq!(rect.y, 0);
        assert_eq!(rect.width, 640);
        assert_eq!(rect.height, 480);
    }

    #[test]
    fn normalize_relative_rect_zero_size_extends_to_edge() {
        let requested = RelativeRect {
            x: 5,
            y: 7,
            width: 0,
            height: 0,
        };
        let rect = normalize_relative_rect(Some(requested), 20, 30).expect("rect should normalize");
        assert_eq!(rect.x, 5);
        assert_eq!(rect.y, 7);
        assert_eq!(rect.width, 15);
        assert_eq!(rect.height, 23);
    }

    #[test]
    fn normalize_relative_rect_out_of_bounds_fails() {
        let requested = RelativeRect {
            x: 8,
            y: 8,
            width: 5,
            height: 5,
        };
        let err = normalize_relative_rect(Some(requested), 10, 10)
            .expect_err("out-of-bounds crop should fail");
        assert!(err.to_string().contains("requested crop out of bounds"));
    }

    #[test]
    fn crop_rgba_extracts_expected_region() {
        let source: Vec<u8> = vec![
            1, 2, 3, 4, 11, 12, 13, 14, 21, 22, 23, 24, // row 0
            31, 32, 33, 34, 41, 42, 43, 44, 51, 52, 53, 54, // row 1
        ];
        let rect = RelativeRect {
            x: 1,
            y: 0,
            width: 2,
            height: 2,
        };
        let cropped = crop_rgba(&source, 3, 2, rect);
        let expected = vec![
            11, 12, 13, 14, 21, 22, 23, 24, // from row 0
            41, 42, 43, 44, 51, 52, 53, 54, // from row 1
        ];
        assert_eq!(cropped, expected);
    }

    #[test]
    fn resolve_output_path_for_new_directory_returns_png_file() {
        let root = std::env::temp_dir().join(format!("capture-mcp-server-test-{}", Uuid::new_v4()));
        let result = resolve_output_path(root.to_string_lossy().as_ref())
            .expect("directory input should resolve to a file path");

        assert!(result.starts_with(&root));
        assert_eq!(result.extension().and_then(|s| s.to_str()), Some("png"));
        assert!(root.exists());

        if root.exists() {
            fs::remove_dir_all(&root).expect("temp directory cleanup should succeed");
        }
    }
}
