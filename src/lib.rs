use image::codecs::jpeg::JpegEncoder;
use windows::Win32::System::Com::COINIT_APARTMENTTHREADED;
use windows::Win32::System::Com::COINIT_DISABLE_OLE1DDE;
// use std::ffi::c_void;
use std::ffi::OsStr;
use std::io::Cursor;
use std::mem::size_of;
use std::os::windows::ffi::OsStrExt;
use windows::core::{Interface, PCWSTR};
use windows::Win32::Foundation::SIZE;
use windows::Win32::Graphics::Gdi::{
    CreateCompatibleDC, DeleteDC, DeleteObject, GetDIBits, GetObjectW, BITMAP, BITMAPFILEHEADER,
    BITMAPINFO, BITMAPINFOHEADER, BI_RGB, DIB_RGB_COLORS, HGDIOBJ,
};
use windows::Win32::System::Com::{CoInitializeEx, CoUninitialize, COINIT};
use windows::Win32::UI::Shell::{
    IShellItem, IShellItemImageFactory, SHCreateItemFromParsingName, SIIGBF_THUMBNAILONLY,
};

// COM初始化（每个线程只初始化一次）
thread_local! {
    static COM_LIB: ComLibrary = ComLibrary::init().expect("COM init failed");
}

// RAII wrapper for COM initialization
struct ComLibrary;
impl ComLibrary {
    fn init() -> Result<Self, windows::core::Error> {
        unsafe {
            let _hr = CoInitializeEx(
                None,
                COINIT(COINIT_APARTMENTTHREADED.0 | COINIT_DISABLE_OLE1DDE.0),
            );
            Ok(Self)
        }
    }
}

impl Drop for ComLibrary {
    fn drop(&mut self) {
        unsafe {
            CoUninitialize();
        }
    }
}

// RAII wrapper for HDC
struct SafeDc(windows::Win32::Graphics::Gdi::HDC);
impl Drop for SafeDc {
    fn drop(&mut self) {
        unsafe {
            let _ = DeleteDC(self.0);
        }
    }
}

// RAII wrapper for GDI objects
struct SafeGdiObj(HGDIOBJ);
impl Drop for SafeGdiObj {
    fn drop(&mut self) {
        unsafe {
            let _ = DeleteObject(self.0);
        }
    }
}

fn create_shell_item(file_name: &str) -> Result<IShellItem, windows::core::Error> {
    let wide_file_name: Vec<u16> = OsStr::new(file_name).encode_wide().chain(Some(0)).collect();
    unsafe { SHCreateItemFromParsingName(PCWSTR(wide_file_name.as_ptr()), None) }
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
        match self {
            Self::S16 => SIZE { cx: 16, cy: 16 },
            Self::S32 => SIZE { cx: 32, cy: 32 },
            Self::S48 => SIZE { cx: 48, cy: 48 },
            Self::S96 => SIZE { cx: 96, cy: 96 },
            Self::S256 => SIZE { cx: 256, cy: 256 },
            Self::S768 => SIZE { cx: 768, cy: 768 },
            Self::S1280 => SIZE { cx: 1280, cy: 1280 },
            Self::S1920 => SIZE { cx: 1920, cy: 1920 },
            Self::S2560 => SIZE { cx: 2560, cy: 2560 },
            Self::Custom(w, h) => SIZE { cx: w, cy: h },
        }
    }
}

/// Returns thumbnail bitmap bits
/// Thumbnail will be no larger than the specified width and height.
///
/// ```
/// let bmp = thumbcache::get_bmp(r"C:\path-to-file.jpeg", thumbcache::ThumbSize::S96).unwrap()
/// ```
pub fn get_bmp(file_path: &str, size: ThumbSize) -> Result<Vec<u8>, windows::core::Error> {
    get_bits(file_path, size)
}

/// 获取指定文件路径的缩略图JPEG数据。
///
/// 该函数通过调用 `get_jpeg_with_quality` 函数生成指定大小的缩略图，并使用默认的JPEG质量（85）进行编码。
///
/// # 参数
/// - `file_path`: 字符串切片，表示要生成缩略图的文件路径。
/// - `size`: `ThumbSize` 类型，表示缩略图的目标尺寸。
///
/// # 示例
/// ```
/// let jpeg_data = thumbcache::get_jpeg(r"C:\path-to-file.jpeg", thumbcache::ThumbSize::S96).unwrap();
/// ```
/// # 返回值
/// - `Result<Vec<u8>, windows::core::Error>`: 如果成功，返回包含JPEG数据的字节向量；如果失败，返回 `windows::core::Error` 类型的错误。
pub fn get_jpeg(file_path: &str, size: ThumbSize) -> Result<Vec<u8>, windows::core::Error> {
    get_jpeg_with_quality(file_path, size, 85)
}

/// 根据指定的文件路径、缩略图大小和质量，生成自定义质量的JPEG图像。
///
/// # 参数
/// - `file_path`: 字符串切片，表示要处理的图像文件的路径。
/// - `size`: `ThumbSize` 类型，表示生成的缩略图的大小。
/// - `quality`: `u8` 类型，表示生成的JPEG图像的质量，范围为0到100，100为最高质量。
///
/// # 示例
/// ```
/// let jpeg_data = thumbcache::get_jpeg_with_quality(r"C:\path-to-file.jpeg", thumbcache::ThumbSize::S96, 85).unwrap();
/// ```
///
/// # 返回值
/// - `Result<Vec<u8>, windows::core::Error>`: 如果成功，返回包含JPEG图像数据的字节向量；如果失败，返回 `windows::core::Error` 错误。
///
/// # 错误
/// - 如果无法从文件路径读取位图数据，或者无法将位图数据加载为图像，或者无法将图像编码为JPEG格式，函数将返回相应的错误。
pub fn get_jpeg_with_quality(
    file_path: &str,
    size: ThumbSize,
    quality: u8,
) -> Result<Vec<u8>, windows::core::Error> {
    let bmp_bytes = get_bits(file_path, size)?;

    let img = image::load_from_memory(&bmp_bytes).map_err(|e| {
        windows::core::Error::new(windows::core::HRESULT::from_win32(1), e.to_string())
    })?;

    let mut jpeg_data = Vec::new();
    let mut cursor = Cursor::new(&mut jpeg_data);

    JpegEncoder::new_with_quality(&mut cursor, quality)
        .encode_image(&img)
        .map_err(|e| {
            windows::core::Error::new(windows::core::HRESULT::from_win32(1), e.to_string())
        })?;

    Ok(jpeg_data)
}

fn get_bits(file_path: &str, size: ThumbSize) -> Result<Vec<u8>, windows::core::Error> {
    COM_LIB.with(|_| {});
    let hbitmap = unsafe {
        let shell_item = create_shell_item(file_path)?;
        let factory: IShellItemImageFactory = shell_item.cast()?;
        factory.GetImage(size.to_size(), SIIGBF_THUMBNAILONLY)?
    };

    let hgdiobj = SafeGdiObj(hbitmap.into());

    unsafe {
        let mut bmp = BITMAP::default();
        GetObjectW(
            hgdiobj.0,
            size_of::<BITMAP>() as i32,
            Some(&mut bmp as *mut _ as *mut _),
        );

        let hdc = SafeDc(CreateCompatibleDC(None));

        let mut bitmap_info = BITMAPINFO {
            bmiHeader: BITMAPINFOHEADER {
                biSize: size_of::<BITMAPINFOHEADER>() as u32,
                biWidth: bmp.bmWidth,
                biHeight: -bmp.bmHeight, // 使用负高度表示top-down DIB
                biPlanes: 1,
                biBitCount: 32, // 强制32位格式避免调色板问题
                biCompression: BI_RGB.0,
                ..Default::default()
            },
            ..Default::default()
        };

        let byte_size = (bmp.bmWidth * 4).abs() * bmp.bmHeight.abs();
        let mut bits = vec![0u8; byte_size as usize];

        if GetDIBits(
            hdc.0,
            hbitmap,
            0,
            bmp.bmHeight.unsigned_abs(),
            Some(bits.as_mut_ptr() as _),
            &mut bitmap_info,
            DIB_RGB_COLORS,
        ) == 0
        {
            return Err(windows::core::Error::from_win32());
        }

        // 构造BMP文件
        let file_header = BITMAPFILEHEADER {
            bfType: 0x4D42,
            bfSize: (size_of::<BITMAPFILEHEADER>() + size_of::<BITMAPINFOHEADER>() + bits.len())
                as u32,
            bfOffBits: (size_of::<BITMAPFILEHEADER>() + size_of::<BITMAPINFOHEADER>()) as u32,
            ..Default::default()
        };

        let mut result = Vec::with_capacity(
            size_of::<BITMAPFILEHEADER>() + size_of::<BITMAPINFOHEADER>() + bits.len(),
        );

        result.extend_from_slice(std::slice::from_raw_parts(
            &file_header as *const _ as *const u8,
            size_of::<BITMAPFILEHEADER>(),
        ));

        result.extend_from_slice(std::slice::from_raw_parts(
            &bitmap_info.bmiHeader as *const _ as *const u8,
            size_of::<BITMAPINFOHEADER>(),
        ));

        result.extend_from_slice(&bits);

        Ok(result)
    }
}
