use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

use crate::macros::wrap_struct;

wrap_struct!(
    Formula,
    xlsx::Formula,
    set_result(result: &str)
);
