use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

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
}
