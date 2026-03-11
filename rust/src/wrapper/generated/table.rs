use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `Table` struct represents a worksheet table.
///
/// Tables in Excel are a way of grouping a range of cells into a single entity
/// that has common formatting or that can be referenced in formulas. Tables
/// can have column headers, autofilters, total rows, column formulas, and
/// different formatting styles.
///
/// The image below shows a default table in Excel with the default properties
/// displayed in the ribbon bar.
///
/// A table is added to a worksheet via the
/// {@link Worksheet#addTable} method. The headers
/// and total row of a table should be configured via a `Table` struct, but the
/// table data can be added using standard
/// {@link Worksheet#write} methods:
///
/// Output file:
///
/// For more information on tables see the Microsoft documentation on [Overview
/// of Excel tables].
///
/// [Overview of Excel tables]:
///     https://support.microsoft.com/en-us/office/overview-of-excel-tables-7ab0bb7d-3a9e-4b56-a3c9-6c94334e492c
#[derive(Clone)]
#[wasm_bindgen]
pub struct Table {
    pub(crate) inner: Arc<Mutex<xlsx::Table>>,
}

#[wasm_bindgen]
impl Table {
    #[wasm_bindgen(constructor)]
    pub fn new() -> Table {
        Table {
            inner: Arc::new(Mutex::new(xlsx::Table::new())),
        }
    }
    /// Check if the table has a header row.
    ///
    /// This method is mainly used by polars_excel_writer and hidden from the
    /// general documentation.
    #[wasm_bindgen(js_name = "hasHeaderRow", skip_jsdoc)]
    pub fn has_header_row(&self) -> bool {
        let lock = self.inner.lock().unwrap();
        lock.has_header_row()
    }
    /// Check if the table has a totals row.
    ///
    /// This method is mainly used by polars_excel_writer and hidden from the
    /// general documentation.
    #[wasm_bindgen(js_name = "hasTotalRow", skip_jsdoc)]
    pub fn has_total_row(&self) -> bool {
        let lock = self.inner.lock().unwrap();
        lock.has_total_row()
    }
}
