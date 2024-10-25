use core::fmt;
use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

#[derive(Debug)]
pub enum XlsxError {
    Xlsx(xlsx::XlsxError),
    Type(String),
    Internal(String),
    InvalidDate,
}

impl std::error::Error for XlsxError {}

impl fmt::Display for XlsxError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            XlsxError::Xlsx(e) => write!(f, "XlsxError({:?})", e),
            XlsxError::Type(e) => write!(f, "TypeError({:?})", e),
            XlsxError::Internal(e) => write!(f, "InternalError({:?})", e),
            XlsxError::InvalidDate => write!(f, "InvalidDateError"),
        }
    }
}

impl From<xlsx::XlsxError> for XlsxError {
    fn from(e: xlsx::XlsxError) -> Self {
        XlsxError::Xlsx(e)
    }
}

impl From<XlsxError> for JsValue {
    fn from(e: XlsxError) -> JsValue {
        JsValue::from_str(e.to_string().as_str())
    }
}
