use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct FilterData {
    pub(crate) inner: Arc<Mutex<xlsx::FilterData>>,
}

#[wasm_bindgen]
impl FilterData {}
