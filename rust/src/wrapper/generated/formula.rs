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
    pub fn new(formula: impl AsRef) -> Formula {
        Formula {
            inner: Arc::new(Mutex::new(xlsx::Formula::new(formula))),
        }
    }
}
