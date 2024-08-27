# thumbcache

Uses Windows thumbcache to get bmp preview for a file.

## Usage

When trying to get preview from file that is not an image (.zip or .exe) will throw error. May also throw error if preview does not exists. So it's better to check for errors rather than relying on `.unwrap()` or `?` syntax.

```rs
use std::io::{Error, Write};

pub fn main() -> Result<(), Error> {
  let hbitmap = thumbcache::get_hbitmap(r"C:\path-to-file.jpeg", 96, 96, 0x08)?;
  
  // true  — include .bpm file headers
  // false — dont
  let bitmap = thumbcache::get_bitmap_bits(hbitmap, true);

  let mut file_out = std::fs::File::create("./out.bmp")?;
  let _ = file_out.write_all(&bitmap);
  
  Ok(())
}
```

## Sources

When I was writing this thing I used sources below. Also, I have used ChatGPT because I am not a Rust developer in first place and not very familiar with C or WinAPI. Worth mentioning in my opinion.

https://stackoverflow.com/questions/14207618/get-bytes-from-hbitmap

https://stackoverflow.com/questions/21751747/extract-thumbnail-for-any-file-in-windows