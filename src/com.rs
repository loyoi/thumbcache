use windows::Win32::System::Com::{
    CoInitializeEx, CoUninitialize, COINIT, COINIT_DISABLE_OLE1DDE, COINIT_MULTITHREADED,
};

pub struct ComLibrary;

impl ComLibrary {
    pub fn init() -> ComLibrary {
        unsafe {
            let _ = CoInitializeEx(
                None,
                COINIT(COINIT_MULTITHREADED.0 | COINIT_DISABLE_OLE1DDE.0),
            );

            Self
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
