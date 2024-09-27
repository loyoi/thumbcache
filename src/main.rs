use std::io::{Error, Write};

pub fn main() -> Result<(), Error> {
  let bmp = thumbcache::get_bmp(r"C:\path-to-file.jpeg", 96, 96)?;

  let mut file_out = std::fs::File::create("./out.bmp")?;
  let _ = file_out.write_all(&bmp);
  
  Ok(())
}