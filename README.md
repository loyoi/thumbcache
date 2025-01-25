# thumbcache

Uses Windows thumbcache to get bmp preview for a file.

## Usage

When trying to get preview from file that is not an image (.zip or .exe) will gibe error. May also give error if preview does not exists.

```rs
use std::io::{Error, Write};

pub fn main() -> Result<(), Error> {
  let bmp = thumbcache::get_bmp(r"C:\path-to-file.jpeg", 96, 96)?;
  
  let mut file_out = std::fs::File::create("./out.bmp")?;
  let _ = file_out.write_all(&bmp);
  
  Ok(())
}
```

## Sources

https://stackoverflow.com/questions/14207618/get-bytes-from-hbitmap

https://stackoverflow.com/questions/21751747/extract-thumbnail-for-any-file-in-windows