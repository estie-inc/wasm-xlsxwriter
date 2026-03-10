use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct CustomProperty {
    pub(crate) inner: Arc<Mutex<xlsx::CustomProperty>>,
}

#[wasm_bindgen]
impl CustomProperty {}
