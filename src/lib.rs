use std::ffi::c_void;
use std::mem::{size_of, transmute};
use std::ffi::OsStr;
use std::os::windows::ffi::OsStrExt;
use windows::core::{Interface, PCWSTR};
use windows::Win32::Foundation::SIZE;
use windows::Win32::Graphics::Gdi::{DeleteObject, GetObjectW, GetDIBits, CreateCompatibleDC, DeleteDC, BITMAP, BITMAPFILEHEADER, BITMAPINFO, BITMAPINFOHEADER, BI_RGB, DIB_RGB_COLORS, HDC};
use windows::Win32::System::Com::{CoInitializeEx, COINIT};
use windows::Win32::UI::Shell::{IShellItem, IShellItemImageFactory, SHCreateItemFromParsingName, SIIGBF};

fn create_shell_item(file_name: &str) -> Result<IShellItem, windows::core::Error> {
    unsafe {
        let wide_file_name: Vec<u16> = OsStr::new(file_name).encode_wide().chain(Some(0)).collect();
        let item: Result<IShellItem, windows::core::Error> = SHCreateItemFromParsingName(PCWSTR(wide_file_name.as_ptr()), None);

        return item;
    }
}

/// Returns bitmap bits
/// 
/// ```
/// let bmp = thumbcache::get_bmp(r"C:\path-to-file.jpeg", 96, 96)?
/// ```
pub fn get_bmp(file_path: &str, width: i32, height: i32) -> Result<Vec<u8>, windows::core::Error> {
    let hbitmap = unsafe {
        let _ = CoInitializeEx(None, COINIT(0));

        let shell_item = create_shell_item(file_path)?;
        let factory: IShellItemImageFactory = shell_item.cast()?;

        let size = SIZE {
            cx: width,
            cy: height
        };

        factory.GetImage(size, SIIGBF(0x08))?
    };

    let bytes = unsafe {
        let mut bmp: BITMAP = std::mem::zeroed();
        GetObjectW(hbitmap, std::mem::size_of::<BITMAP>() as i32, Some(&mut bmp as *mut _ as *mut _));

        let hdc = CreateCompatibleDC(HDC(0 as *mut c_void));
        let mut bitmap_info = BITMAPINFO {
            bmiHeader: BITMAPINFOHEADER {
                biSize: std::mem::size_of::<BITMAPINFOHEADER>() as u32,
                biWidth: bmp.bmWidth,
                biHeight: bmp.bmHeight,
                biPlanes: 1,
                biBitCount: bmp.bmBitsPixel,
                biCompression: BI_RGB.0,
                ..Default::default()
            },
            ..Default::default()
        };

        let byte_size = bmp.bmWidthBytes * bmp.bmHeight;
        let mut bits = vec![0u8; byte_size as usize];

        GetDIBits(hdc, hbitmap, 0, bmp.bmHeight as u32, Some(bits.as_mut_ptr() as *mut c_void), &mut bitmap_info, DIB_RGB_COLORS);

        let _ = DeleteDC(hdc);
        let _ = DeleteObject(hbitmap);

        let file_header = BITMAPFILEHEADER {
            bfType: 0x4D42,
            bfSize: (size_of::<BITMAPFILEHEADER>() + size_of::<BITMAPINFOHEADER>() + bits.len()) as u32,
            bfReserved1: 0,
            bfReserved2: 0,
            bfOffBits: (size_of::<BITMAPFILEHEADER>() + size_of::<BITMAPINFOHEADER>()) as u32,
        };

        let file_header_bytes: &[u8] = std::slice::from_raw_parts(transmute(&file_header), size_of::<BITMAPFILEHEADER>());
        let info_header_bytes: &[u8] = std::slice::from_raw_parts(transmute(&bitmap_info), size_of::<BITMAPINFOHEADER>());

        [file_header_bytes, info_header_bytes, &bits].concat()
    };

    return Ok(bytes);
}