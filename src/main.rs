use std::io::{Error, Write};

pub fn main() -> Result<(), Error> {
  let hbitmap = thumbcache::get_hbitmap(r"C:\path-to-file.jpeg", 96, 96, 0x08)?;
  let bitmap = thumbcache::get_bitmap_bits(hbitmap, true);

  let mut file_out = std::fs::File::create("./out.bmp")?;
  let _ = file_out.write_all(&bitmap);
  
  Ok(())
}