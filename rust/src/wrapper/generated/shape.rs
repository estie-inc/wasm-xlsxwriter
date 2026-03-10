use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct Shape {
    pub(crate) inner: Arc<Mutex<xlsx::Shape>>,
}

#[wasm_bindgen]
impl Shape {}
