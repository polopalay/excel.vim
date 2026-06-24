use std::fmt;

/// Lỗi nghiệp vụ tương đương `RuntimeError` trong bản Python.
/// In ra stderr và exit code != 0 khi gặp lỗi, giống hành vi của
/// `echoerr join(...)` trong excelPlugin.vim (đọc output qua v:shell_error).
#[derive(Debug)]
pub struct AppError(pub String);

impl fmt::Display for AppError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.0)
    }
}

impl std::error::Error for AppError {}

impl From<std::io::Error> for AppError {
    fn from(e: std::io::Error) -> Self {
        AppError(format!("IO error: {e}"))
    }
}

impl From<zip::result::ZipError> for AppError {
    fn from(e: zip::result::ZipError) -> Self {
        AppError(format!("Zip error: {e}"))
    }
}

impl From<quick_xml::Error> for AppError {
    fn from(e: quick_xml::Error) -> Self {
        AppError(format!("XML error: {e}"))
    }
}

impl From<std::string::FromUtf8Error> for AppError {
    fn from(e: std::string::FromUtf8Error) -> Self {
        AppError(format!("UTF-8 error: {e}"))
    }
}

impl From<quick_xml::events::attributes::AttrError> for AppError {
    fn from(e: quick_xml::events::attributes::AttrError) -> Self {
        AppError(format!("XML attribute error: {e}"))
    }
}

pub type AppResult<T> = Result<T, AppError>;

/// Helper tạo lỗi nhanh, giống `raise RuntimeError(f"...")` trong Python.
#[macro_export]
macro_rules! bail {
    ($($arg:tt)*) => {
        return Err($crate::error::AppError(format!($($arg)*)))
    };
}
