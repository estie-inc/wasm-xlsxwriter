use rust_xlsxwriter::{self as xlsx};
use wasm_bindgen::prelude::*;

use crate::wrapper::Format;

use super::{Formula, TableStyle};

/// The `Table` struct represents a worksheet Table.
///
/// Tables in Excel are a way of grouping a range of cells into a single entity
/// that has common formatting or that can be referenced from formulas. Tables
/// can have column headers, autofilters, total rows, column formulas and
/// different formatting styles.
///
/// The image below shows a default table in Excel with the default properties
/// shown in the ribbon bar.
///
/// <img src="https://rustxlsxwriter.github.io/images/table_intro.png">
///
/// A table is added to a worksheet via the
/// {@link Worksheet#addTable}(crate::Worksheet::add_table) method. The headers
/// and total row of a table should be configured via a `Table` struct but the
/// table data can be added via standard
/// {@link Worksheet#write}(crate::Worksheet::write) methods:
///
/// TODO: example omitted
///
/// For more information on tables see the Microsoft documentation on [Overview
/// of Excel tables].
///
/// [Overview of Excel tables]:
///     https://support.microsoft.com/en-us/office/overview-of-excel-tables-7ab0bb7d-3a9e-4b56-a3c9-6c94334e492c
#[derive(Clone)]
#[wasm_bindgen]
pub struct Table {
    pub(crate) inner: xlsx::Table,
}

#[wasm_bindgen]
impl Table {
    #[wasm_bindgen(constructor)]
    pub fn new() -> Self {
        Table {
            inner: xlsx::Table::new(),
        }
    }

    #[wasm_bindgen(js_name = "setName")]
    pub fn set_name(&self, name: &str) -> Table {
        Table {
            inner: self.clone().inner.set_name(name),
        }
    }

    #[wasm_bindgen(js_name = "setStyle")]
    pub fn set_style(&self, style: TableStyle) -> Table {
        let style = xlsx::TableStyle::from(style);
        Table {
            inner: self.clone().inner.set_style(style),
        }
    }

    // FIXME: ownership?
    #[wasm_bindgen(js_name = "setColumns")]
    pub fn set_columns(&self, columns: Vec<TableColumn>) -> Table {
        let columns: Vec<_> = columns.into_iter().map(|c| c.inner).collect();
        Table {
            inner: self.clone().inner.set_columns(&columns),
        }
    }

    #[wasm_bindgen(js_name = "setFirstColumn")]
    pub fn set_first_column(&self, enable: bool) -> Table {
        Table {
            inner: self.clone().inner.set_first_column(enable),
        }
    }

    #[wasm_bindgen(js_name = "setHeaderRow")]
    pub fn set_header_row(&self, enable: bool) -> Table {
        Table {
            inner: self.clone().inner.set_header_row(enable),
        }
    }

    #[wasm_bindgen(js_name = "setTotalRow")]
    pub fn set_total_row(&self, enable: bool) -> Table {
        Table {
            inner: self.clone().inner.set_total_row(enable),
        }
    }

    #[wasm_bindgen(js_name = "setBandedColumns")]
    pub fn set_banded_columns(&self, enable: bool) -> Table {
        Table {
            inner: self.clone().inner.set_banded_columns(enable),
        }
    }

    #[wasm_bindgen(js_name = "setBandedRows")]
    pub fn set_banded_rows(&self, enable: bool) -> Table {
        Table {
            inner: self.clone().inner.set_banded_rows(enable),
        }
    }
}

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
    inner: xlsx::TableColumn,
}

#[wasm_bindgen]
impl TableColumn {
    /// Create a new `TableColumn` to configure a Table column.
    #[wasm_bindgen(constructor)]
    pub fn new() -> Self {
        TableColumn {
            inner: xlsx::TableColumn::new(),
        }
    }

    /// Set the header caption for a table column.
    ///
    /// Excel uses default captions such as `Column 1`, `Column 2`, etc., for
    /// the headers on a worksheet table. These can be set to a user defined
    /// value using the `setHeader()` method.
    ///
    /// The column header names in a table must be different from each other.
    /// Non-unique names will raise a validation error when using
    /// {@link Worksheet#addTable}.
    ///
    /// @param {string} caption - The caption/name of the column header. It must be unique for the table.
    /// @returns {TableColumn} - The TableColumn object.
    #[wasm_bindgen(js_name = "setHeader", skip_jsdoc)]
    pub fn set_header(&self, caption: &str) -> TableColumn {
        TableColumn {
            inner: self.clone().inner.set_header(caption),
        }
    }

    /// Set the format for the header of the table column.
    ///
    /// The `setHeaderFormat` method can be used to set the format for the
    /// column header in a worksheet table.
    ///
    /// @param {Format} format - The {@link Format} property for the column header.
    /// @returns {TableColumn} - The TableColumn object.
    #[wasm_bindgen(js_name = "setHeaderFormat", skip_jsdoc)]
    pub fn set_header_format(&self, format: &Format) -> TableColumn {
        TableColumn {
            inner: self.clone().inner.set_header_format(format.inner.lock().unwrap().clone()),
        }
    }

    /// Set the format for a table column.
    ///
    /// It is sometimes required to format the data in the columns of a table.
    /// This can be done using the standard
    /// {@link Worksheet#writeWithFormat} method
    /// but format can also be applied separately using
    /// `TableColumn.setFormat()`.
    ///
    /// The most common format property to set for a table column is the number
    /// format({@link Format#setNumFormat}), see the example below.
    /// TODO: example omitted
    ///
    /// @param {Format} format - The {@link Format} property for the column header.
    /// @returns {TableColumn} - The TableColumn object.
    #[wasm_bindgen(js_name = "setFormat", skip_jsdoc)]
    pub fn set_format(&self, format: &Format) -> TableColumn {
        TableColumn {
            inner: self.clone().inner.set_format(format.inner.lock().unwrap().clone()),
        }
    }

    /// Set the formula for a table column.
    ///
    /// It is a common use case to add a summation column as the last column in a
    /// table. These are constructed with a special class of Excel formulas
    /// called [Structured References] which can refer to an entire table or
    /// rows or columns of data within the table. For example to sum the data
    /// for several columns in a single row might you might use a formula like
    /// this: `SUM(Table1[@[Quarter 1]:[Quarter 4]])`.
    ///
    /// [Structured References]:
    ///     https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
    ///
    /// @param {Formula} formula - The formula to be applied to the column.
    /// @returns {TableColumn} - The TableColumn object.
    #[wasm_bindgen(js_name = "setFormula", skip_jsdoc)]
    pub fn set_formula(&self, formula: &Formula) -> TableColumn {
        TableColumn {
            inner: self.clone().inner.set_formula(&*formula.inner.lock().unwrap()),
        }
    }

    /// Set a label for the total row of a table column.
    ///
    /// It is possible to set a label for the totals row of a column instead of
    /// a subtotal function. This is most often used to set a caption like
    /// "Totals", as in the example above.
    ///
    /// Note, overwriting the total row cells with `worksheet.write()` calls
    /// will cause Excel to warn that the table is corrupt when loading the
    /// file.
    ///
    /// @param {string} label - The label/caption of the total row of the column.
    /// @returns {TableColumn} - The TableColumn object.
    #[wasm_bindgen(js_name = "setTotalLabel", skip_jsdoc)]
    pub fn set_total_label(&self, label: &str) -> TableColumn {
        TableColumn {
            inner: self.clone().inner.set_total_label(label),
        }
    }

    /// Set the total function for the total row of a table column.
    ///
    /// Set the `SUBTOTAL()` function for the "totals" row of a table column.
    ///
    /// The standard Excel subtotal functions are available via the
    /// {@link TableFunction} enum values. The Excel functions are:
    ///
    /// - Average
    /// - Count
    /// - Count Numbers
    /// - Maximum
    /// - Minimum
    /// - Sum
    /// - Standard Deviation
    /// - Variance
    /// - Custom - User defined function or formula
    ///
    /// Note, overwriting the total row cells with `worksheet.write()` calls
    /// will cause Excel to warn that the table is corrupt when loading the
    /// file.
    ///
    /// @param {TableFunction} table_function - A {@link TableFunction} enum value equivalent to one of the
    ///   available Excel `SUBTOTAL()` options.
    /// @returns {TableColumn} - The TableColumn object.
    #[wasm_bindgen(js_name = "setTotalFunction", skip_jsdoc)]
    pub fn set_total_function(&self, table_function: &TableFunction) -> TableColumn {
        TableColumn {
            inner: self
                .clone()
                .inner
                .set_total_function(table_function.inner.clone()),
        }
    }
}

/// The `TableFunction` enum defines functions for worksheet table total rows.
///
/// The `TableFunction` enum contains definitions for the standard Excel
/// "SUBTOTAL" functions that are available via the dropdown in the total row of
/// an Excel table. It also supports custom user defined functions or formulas.
///
/// TODO: example omitted
#[wasm_bindgen]
pub struct TableFunction {
    inner: xlsx::TableFunction,
}

#[wasm_bindgen]
impl TableFunction {
    /// Use the average function as the table total.
    #[wasm_bindgen]
    pub fn average() -> TableFunction {
        TableFunction {
            inner: xlsx::TableFunction::Average,
        }
    }

    /// Use the count function as the table total.
    #[wasm_bindgen]
    pub fn count() -> TableFunction {
        TableFunction {
            inner: xlsx::TableFunction::Count,
        }
    }

    /// Use the count numbers function as the table total.
    #[wasm_bindgen(js_name = "countNumbers")]
    pub fn count_numbers() -> TableFunction {
        TableFunction {
            inner: xlsx::TableFunction::CountNumbers,
        }
    }

    /// Use the max function as the table total.
    #[wasm_bindgen]
    pub fn max() -> TableFunction {
        TableFunction {
            inner: xlsx::TableFunction::Max,
        }
    }

    /// Use the min function as the table total.
    #[wasm_bindgen]
    pub fn min() -> TableFunction {
        TableFunction {
            inner: xlsx::TableFunction::Min,
        }
    }

    /// Use the sum function as the table total.
    #[wasm_bindgen]
    pub fn sum() -> TableFunction {
        TableFunction {
            inner: xlsx::TableFunction::Sum,
        }
    }

    /// Use the standard deviation function as the table total.
    #[wasm_bindgen(js_name = "stdDev")]
    pub fn std_dev() -> TableFunction {
        TableFunction {
            inner: xlsx::TableFunction::StdDev,
        }
    }

    /// Use the var function as the table total.
    #[wasm_bindgen]
    pub fn var() -> TableFunction {
        TableFunction {
            inner: xlsx::TableFunction::Var,
        }
    }

    /// Use a custom/user specified function or formula.
    #[wasm_bindgen]
    pub fn custom(formula: &Formula) -> TableFunction {
        TableFunction {
            inner: xlsx::TableFunction::Custom(formula.inner.lock().unwrap().clone()),
        }
    }
}
