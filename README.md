# thumbcache

Uses Windows thumbcache to get bmp preview for a file.

## Usage

When trying to get preview from file that is not an image (.zip or .exe) will result in error.

```rs
use std::io::{Error, Write};

pub fn main() -> Result<(), Error> {
  let bmp = thumbcache::get_bmp(r"C:\path-to-file.jpeg", thumbcache::ThumbSize::S96)?;
  
  let mut file_out = std::fs::File::create("./out.bmp")?;
  let _ = file_out.write_all(&bmp);
  
  Ok(())
}
```

## Examples

### Convert BMP into JPEG

Windows can only return BMP, but this format may not always be convenient. In this example, BMP is converted to JPEG using [image](https://github.com/image-rs/image) crate.

```rs
use thumbcache::{get_bmp};
use image::{load_from_memory};

fn main() {
  let bmp = get_bmp(r"C:\path-to-file.jpeg", thumbcache::ThumbSize::S256).unwrap();

  let image = load_from_memory(&bmp).unwrap();

  let mut buf: Vec<u8> = Vec::new();
  let mut writer = std::io::Cursor::new(&mut buf);

  image.write_to(&mut writer, image::ImageFormat::Jpeg).unwrap();

  std::fs::write("./output.jpeg", buf).unwrap();
}
```

## Sources

https://stackoverflow.com/questions/14207618/get-bytes-from-hbitmap

https://stackoverflow.com/questions/21751747/extract-thumbnail-for-any-file-in-windows

https://learn.microsoft.com/en-us/windows/win32/api/shobjidl_core/nf-shobjidl_core-ishellitemimagefactory-getimage