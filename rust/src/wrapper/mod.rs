mod chart;
mod color;
mod doc_properties;
mod excel_data;
mod format;
mod formula;
mod image;
mod note;
mod rich_string;
mod table;
mod url;
mod utils;
mod workbook;
mod worksheet;

use crate::error::XlsxError;

type WasmResult<T> = std::result::Result<T, XlsxError>;
