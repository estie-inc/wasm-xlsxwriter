use crate::wrapper::Format;
use crate::wrapper::Formula;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `TableColumn` struct represents a table column.
///
/// The `TableColumn` struct is used to set the properties for columns in a
/// worksheet {@link Table}. This can be used to set the following properties of a
/// table column:
///
/// - The header caption.
/// - The total row caption.
/// - The total row subtotal function.
/// - A formula for the column.
///
/// This struct is used in conjunction with the {@link Table#setColumns} method.
#[derive(Clone)]
#[wasm_bindgen]
pub struct TableColumn {
    pub(crate) inner: Arc<Mutex<xlsx::TableColumn>>,
}

#[wasm_bindgen]
impl TableColumn {
    #[wasm_bindgen(constructor)]
    pub fn new() -> TableColumn {
        TableColumn {
            inner: Arc::new(Mutex::new(xlsx::TableColumn::new())),
        }
    }
    /// Set the header caption for a table column.
    ///
    /// Excel uses default captions such as `Column 1`, `Column 2`, etc., for
    /// the headers on a worksheet table. These can be set to a user-defined
    /// value using the `set_header()` method.
    ///
    /// The column header names in a table must be different from each other.
    /// Non-unique names will raise a validation error when using
    /// {@link Worksheet#addTable}.
    ///
    /// # Parameters
    ///
    /// - `caption`: The caption/name of the column header. It must be unique
    ///   for the table.
    #[wasm_bindgen(js_name = "setHeader", skip_jsdoc)]
    pub fn set_header(&self, caption: &str) -> TableColumn {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_header(caption);
        *lock = inner;
        TableColumn {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set a label for the total row of a table column.
    ///
    /// It is possible to set a label for the totals row of a column instead of
    /// a subtotal function. This is most often used to set a caption like
    /// "Totals," as in the example above.
    ///
    /// Note, overwriting the total row cells with `worksheet.write()` calls
    /// will cause Excel to warn that the table is corrupt when loading the
    /// file.
    ///
    /// # Parameters
    ///
    /// - `label`: The label/caption of the total row of the column.
    #[wasm_bindgen(js_name = "setTotalLabel", skip_jsdoc)]
    pub fn set_total_label(&self, label: &str) -> TableColumn {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_total_label(label);
        *lock = inner;
        TableColumn {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the formula for a table column.
    ///
    /// It is a common use case to add a summation column as the last column in a
    /// table. These are constructed with a special class of Excel formulas
    /// called [Structured References], which can refer to an entire table or
    /// rows or columns of data within the table. For example, to sum the data
    /// for several columns in a single row, you might use a formula like
    /// this: `SUM(Table1[@[Quarter 1]:[Quarter 4]])`.
    ///
    /// [Structured References]:
    ///     https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
    ///
    /// # Parameters
    ///
    /// - `formula`: The formula to be applied to the column as a string or
    ///   {@link Formula}.
    #[wasm_bindgen(js_name = "setFormula", skip_jsdoc)]
    pub fn set_formula(&self, formula: Formula) -> TableColumn {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_formula(formula.inner.clone());
        *lock = inner;
        TableColumn {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the format for a table column.
    ///
    /// It is sometimes required to format the data in the columns of a table.
    /// This can be done using the standard
    /// {@link Worksheet#writeWithFormat} method,
    /// but the format can also be applied separately using
    /// `TableColumn.set_format()`.
    ///
    /// The most common format property to set for a table column is the [number
    /// format](Format::set_num_format), see the example below.
    ///
    /// # Parameters
    ///
    /// - `format`: The {@link Format} property for the data cells in the column.
    #[wasm_bindgen(js_name = "setFormat", skip_jsdoc)]
    pub fn set_format(&self, format: Format) -> TableColumn {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_format(format.inner.clone());
        *lock = inner;
        TableColumn {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the format for the header of the table column.
    ///
    /// The `set_header_format` method can be used to set the format for the
    /// column header in a worksheet table.
    ///
    /// # Parameters
    ///
    /// - `format`: The {@link Format} property for the column header.
    #[wasm_bindgen(js_name = "setHeaderFormat", skip_jsdoc)]
    pub fn set_header_format(&self, format: Format) -> TableColumn {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_header_format(format.inner.clone());
        *lock = inner;
        TableColumn {
            inner: Arc::clone(&self.inner),
        }
    }
}
