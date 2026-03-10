use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct Comment {
    pub(crate) inner: Arc<Mutex<xlsx::Comment>>,
}

#[wasm_bindgen]
impl Comment {}
