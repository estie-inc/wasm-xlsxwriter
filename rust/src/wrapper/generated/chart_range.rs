use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartRange {
    pub(crate) inner: Arc<Mutex<xlsx::ChartRange>>,
}

#[wasm_bindgen]
impl ChartRange {}
