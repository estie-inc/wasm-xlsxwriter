use crate::wrapper::Color;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct ShapeGradientStop {
    pub(crate) inner: Arc<Mutex<xlsx::ShapeGradientStop>>,
}

#[wasm_bindgen]
impl ShapeGradientStop {
    #[wasm_bindgen(constructor)]
    pub fn new(color: Color, position: u8) -> ShapeGradientStop {
        ShapeGradientStop {
            inner: Arc::new(Mutex::new(xlsx::ShapeGradientStop::new(
                color.inner,
                position,
            ))),
        }
    }
}
