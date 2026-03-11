use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `Formula` struct is used to define a worksheet formula.
///
/// The `Formula` struct creates a formula type that can be used to write
/// worksheet formulas.
///
/// In general, you would use the
/// {@link Worksheet#writeFormula} with a
/// string representation of the formula, like this:
///
/// The formula will then be displayed as expected in Excel:
///
/// src="https://rustxlsxwriter.github.io/images/working_with_formulas1.png">
///
/// To differentiate a formula from an ordinary string (for example,
/// when storing it in a data structure), you can also represent the formula with
/// a {@link Formula} struct:
///
/// Using a `Formula` struct also allows you to write a formula using the
/// generic {@link Worksheet#write} method:
///
/// As shown in the examples above, you can write a formula and expect to have
/// the result appear in Excel. However, there are a few potential issues and
/// differences that the user of `rust_xlsxwriter` should be aware of. These are
/// explained in the sections below.
#[derive(Clone)]
#[wasm_bindgen]
pub struct Formula {
    pub(crate) inner: Arc<Mutex<xlsx::Formula>>,
}

#[wasm_bindgen]
impl Formula {
    #[wasm_bindgen(constructor)]
    pub fn new(formula: &str) -> Formula {
        Formula {
            inner: Arc::new(Mutex::new(xlsx::Formula::new(formula))),
        }
    }
    /// Specify the result of a formula.
    ///
    /// As explained above in the section on [Formula
    /// Results](#formula-results) it is occasionally necessary to specify the
    /// result of a formula. This can be done using the `set_result()` method.
    ///
    /// # Parameters
    ///
    /// `result` - The formula result, as a string or string like type.
    #[wasm_bindgen(js_name = "setResult", skip_jsdoc)]
    pub fn set_result(&self, result: &str) -> Formula {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Formula::new(""));
        inner = inner.set_result(result);
        *lock = inner;
        Formula {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Escape table functions in the formula.
    ///
    /// This method is used to escape table functions in a formula. This mainly
    /// involves converting Excel 2010 `@` table references to the older 2007
    /// format `[#This Row],`. This is required by Excel for backward
    /// compatibility.
    ///
    /// This function is public but hidden since it is mainly only required by
    /// `polars_excel_writer` or other third party wrappers.
    #[wasm_bindgen(js_name = "escapeTableFunctions", skip_jsdoc)]
    pub fn escape_table_functions(&self) -> Formula {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Formula::new(""));
        inner = inner.escape_table_functions();
        *lock = inner;
        Formula {
            inner: Arc::clone(&self.inner),
        }
    }
}
