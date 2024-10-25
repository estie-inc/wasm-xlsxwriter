mod chart;
mod color;
mod doc_properties;
mod excel_data;
mod format;
mod formula;
mod image;
mod table;
mod url;
mod utils;
mod workbook;
mod worksheet;

use rust_xlsxwriter as xlsx;

use crate::error::XlsxError;

fn map_xlsx_error<T>(e: Result<T, xlsx::XlsxError>) -> Result<T, XlsxError> {
    e.map_err(|e| e.into())
}

type WasmResult<T> = std::result::Result<T, XlsxError>;
