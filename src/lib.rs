use std::ffi::c_void;
use std::mem::{size_of, transmute};
use std::ffi::OsStr;
use std::os::windows::ffi::OsStrExt;
use windows::core::{Interface, PCWSTR};
use windows::Win32::Foundation::SIZE;
use windows::Win32::Graphics::Gdi::{CreateCompatibleDC, DeleteDC, DeleteObject, GetDIBits, GetObjectW, BITMAP, BITMAPFILEHEADER, BITMAPINFO, BITMAPINFOHEADER, BI_RGB, DIB_RGB_COLORS, HGDIOBJ};
use windows::Win32::System::Com::{CoInitializeEx, COINIT};
use windows::Win32::UI::Shell::{IShellItem, IShellItemImageFactory, SHCreateItemFromParsingName, SIIGBF_THUMBNAILONLY};

fn create_shell_item(file_name: &str) -> Result<IShellItem, windows::core::Error> {
    unsafe {
        let wide_file_name: Vec<u16> = OsStr::new(file_name).encode_wide().chain(Some(0)).collect();
        let item: Result<IShellItem, windows::core::Error> = SHCreateItemFromParsingName(PCWSTR(wide_file_name.as_ptr()), None);

        return item;
    }
}

#[derive(Debug)]
pub enum ThumbSize {
    /// 16x16 pixels
    S16,
    /// 32x32 pixels
    S32,
    /// 48x48 pixels
    S48,
    /// 96x96 pixels
    S96,
    /// 256x256 pixels
    S256,
    /// 768x768 pixels
    S768,
    /// 1280x1280 pixels
    S1280,
    /// 1920x1920 pixels
    S1920,
    /// 2560x2560 pixels
    S2560,
    /// Custom thumbnail size as (width, height) in pixels
    Custom(i32, i32),
}

impl ThumbSize {
    pub fn to_size(self) -> SIZE {
        let (cx, cy) = match self {
            ThumbSize::S16 => (16, 16),
            ThumbSize::S32 => (32, 32),
            ThumbSize::S48 => (48, 48),
            ThumbSize::S96 => (96, 96),
            ThumbSize::S256 => (256, 256),
            ThumbSize::S768 => (768, 768),
            ThumbSize::S1280 => (1280, 1280),
            ThumbSize::S1920 => (1920, 1920),
            ThumbSize::S2560 => (2560, 2560),
            ThumbSize::Custom(w, h) => (w, h),
        };

        SIZE {
            cx,
            cy
        }
    }
}

/// Returns thumbnail bitmap bits
/// Thumbnail will be no larger than the specified width and height.
/// 
/// ```
/// let bmp = thumbcache::get_bmp(r"C:\path-to-file.jpeg", thumbcache::ThumbSize::S96)?
/// ```
pub fn get_bmp(file_path: &str, size: ThumbSize) -> Result<Vec<u8>, windows::core::Error> {
    let hbitmap = unsafe {
        let _ = CoInitializeEx(None, COINIT(0));

        let shell_item = create_shell_item(file_path)?;
        let factory: IShellItemImageFactory = shell_item.cast()?;

        factory.GetImage(size.to_size(), SIIGBF_THUMBNAILONLY)?
    };

    let bytes = unsafe {
        let mut bmp: BITMAP = std::mem::zeroed();
        let hgdiobj: HGDIOBJ = hbitmap.into();

        GetObjectW(hgdiobj, std::mem::size_of::<BITMAP>() as i32, Some(&mut bmp as *mut _ as *mut _));

        let hdc = CreateCompatibleDC(None);
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
        let _ = DeleteObject(hgdiobj);

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