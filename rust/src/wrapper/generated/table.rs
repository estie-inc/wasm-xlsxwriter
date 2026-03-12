use crate::wrapper::TableColumn;
use crate::wrapper::TableStyle;
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
    #[doc = r" Create a deep clone of this object."]
    #[wasm_bindgen(js_name = "clone")]
    pub fn deep_clone(&self) -> Table {
        Table {
            inner: Arc::new(Mutex::new(self.inner.lock().unwrap().clone())),
        }
    }
    /// Turn on/off the header row for a table.
    ///
    /// Turn on or off the header row in the table. The header row displays the
    /// column names and, unless it is turned off, an autofilter. It is on by
    /// default.
    ///
    /// The header row will display default captions such as `Column 1`, `Column
    /// 2`, etc. These captions can be overridden using the
    /// {@link Table#setColumns} method, see the examples below. They shouldn't
    /// be written or overwritten using standard
    /// {@link Worksheet#write} methods since that will
    /// cause a warning when the file is loaded in Excel.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is on by default.
    #[wasm_bindgen(js_name = "setHeaderRow", skip_jsdoc)]
    pub fn set_header_row(&self, enable: bool) -> Table {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Table::new());
        inner = inner.set_header_row(enable);
        *lock = inner;
        Table {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Turn on a totals row for a table.
    ///
    /// The `set_total_row()` method can be used to turn on the total row in the
    /// last row of a table. The total row is distinguished from the other rows
    /// by a different formatting and with dropdown `SUBTOTAL()` functions.
    ///
    /// Note, you will need to use {@link TableColumn} methods to populate this row.
    /// Overwriting the total row cells with `worksheet.write()` calls will
    /// cause Excel to warn that the table is corrupt when loading the file.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setTotalRow", skip_jsdoc)]
    pub fn set_total_row(&self, enable: bool) -> Table {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Table::new());
        inner = inner.set_total_row(enable);
        *lock = inner;
        Table {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Turn on/off banded rows for a table.
    ///
    /// By default, Excel uses "banded" rows of alternating colors in a table to
    /// distinguish each data row, like this:
    ///
    /// src="https://rustxlsxwriter.github.io/images/table_set_header_row2.png">
    ///
    /// If you prefer not to have this type of formatting, you can turn it off,
    /// see the example below.
    ///
    /// Note, you can also select a table style without banded rows using the
    /// {@link Table#setStyle} method.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is on by default.
    #[wasm_bindgen(js_name = "setBandedRows", skip_jsdoc)]
    pub fn set_banded_rows(&self, enable: bool) -> Table {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Table::new());
        inner = inner.set_banded_rows(enable);
        *lock = inner;
        Table {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Turn on/off banded columns for a table.
    ///
    /// By default, Excel uses the same format color for each data column in a
    /// table but alternates the color of rows. If you wish, you can set "banded"
    /// columns of alternating colors in a table to distinguish each data column.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setBandedColumns", skip_jsdoc)]
    pub fn set_banded_columns(&self, enable: bool) -> Table {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Table::new());
        inner = inner.set_banded_columns(enable);
        *lock = inner;
        Table {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Turn on/off the first column highlighting for a table.
    ///
    /// The first column of a worksheet table is often used for a list of items,
    /// whereas the other columns are more commonly used for data. In these
    /// cases, it is sometimes desirable to highlight the first column differently.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setFirstColumn", skip_jsdoc)]
    pub fn set_first_column(&self, enable: bool) -> Table {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Table::new());
        inner = inner.set_first_column(enable);
        *lock = inner;
        Table {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Turn on/off the last column highlighting for a table.
    ///
    /// The last column of a worksheet table is often used for a `SUM()` or
    /// other formula operating on the data in the other columns. In these
    /// cases, it is sometimes required to highlight the last column differently.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setLastColumn", skip_jsdoc)]
    pub fn set_last_column(&self, enable: bool) -> Table {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Table::new());
        inner = inner.set_last_column(enable);
        *lock = inner;
        Table {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Turn on/off the autofilter for a table.
    ///
    /// By default, Excel adds an autofilter to the header of a table. This
    /// method can be used to turn it off if necessary.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is on by default.
    #[wasm_bindgen(js_name = "setAutofilter", skip_jsdoc)]
    pub fn set_autofilter(&self, enable: bool) -> Table {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Table::new());
        inner = inner.set_autofilter(enable);
        *lock = inner;
        Table {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set properties for the columns in a table.
    ///
    /// Set the properties for columns in a worksheet table via an array of
    /// {@link TableColumn} structs. This can be used to set the following
    /// properties of a table column:
    ///
    /// - The header caption.
    /// - The total row caption.
    /// - The total row subtotal function.
    /// - A formula for the column.
    ///
    /// # Parameters
    ///
    /// - `columns`: An array reference of {@link TableColumn} structs. Use
    ///   `TableColumn::default()` to get default values.
    #[wasm_bindgen(js_name = "setColumns", skip_jsdoc)]
    pub fn set_columns(&self, columns: Vec<TableColumn>) -> Table {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Table::new());
        inner = inner.set_columns(&columns.iter().map(|x| x.inner.clone()).collect::<Vec<_>>());
        *lock = inner;
        Table {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the name for a table.
    ///
    /// The name of a worksheet table in Excel is similar to a defined name
    /// representing a data region, and it can be used in structured reference
    /// formulas.
    ///
    /// By default, Excel and `rust_xlsxwriter` use a `Table1` .. `TableN`
    /// naming convention for tables in a workbook. If required, you can set a
    /// user-defined name. However, you need to ensure that this name is unique
    /// across the workbook; otherwise, you will get a warning when you load the
    /// file in Excel.
    ///
    /// # Parameters
    ///
    /// - `name`: The name of the table. It must be unique across the workbook.
    #[wasm_bindgen(js_name = "setName", skip_jsdoc)]
    pub fn set_name(&self, name: &str) -> Table {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Table::new());
        inner = inner.set_name(name);
        *lock = inner;
        Table {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the style for a table.
    ///
    /// Excel supports 61 different styles for tables divided into Light, Medium,
    /// and Dark categories.
    ///
    /// You can set one of these styles using a {@link TableStyle} enum value. The
    /// default table style in Excel is equivalent to {@link TableStyle#Medium9}.
    ///
    /// # Parameters
    ///
    /// - `style`: a {@link TableStyle} enum value.
    #[wasm_bindgen(js_name = "setStyle", skip_jsdoc)]
    pub fn set_style(&self, style: TableStyle) -> Table {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Table::new());
        inner = inner.set_style(xlsx::TableStyle::from(style));
        *lock = inner;
        Table {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the alt text for the table to help accessibility.
    ///
    /// The alt text is used with screen readers to help people with visual
    /// disabilities.
    ///
    /// See the following Microsoft documentation on [Everything you need to
    /// know to write effective alt
    /// text](https://support.microsoft.com/en-us/office/everything-you-need-to-know-to-write-effective-alt-text-df98f884-ca3d-456c-807b-1a1fa82f5dc2).
    ///
    /// # Parameters
    ///
    /// - `alt_text`: The alt text string to add to the table.
    #[wasm_bindgen(js_name = "setAltText", skip_jsdoc)]
    pub fn set_alt_text(&self, alt_text: &str) -> Table {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Table::new());
        inner = inner.set_alt_text(alt_text);
        *lock = inner;
        Table {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the alt text title for the table to help accessibility.
    ///
    /// The alt text title is used in conjunction with the alt text, and the
    /// {@link Table#setAltText} method, to provide context for the table. It
    /// is only displayed as a title in Excel for Windows.
    ///
    /// # Parameters
    ///
    /// - `title`: The alt text title string to add to the table.
    #[wasm_bindgen(js_name = "setAltTextTitle", skip_jsdoc)]
    pub fn set_alt_text_title(&self, title: &str) -> Table {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Table::new());
        inner = inner.set_alt_text_title(title);
        *lock = inner;
        Table {
            inner: Arc::clone(&self.inner),
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
