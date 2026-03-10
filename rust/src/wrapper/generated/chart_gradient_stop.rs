use crate::wrapper::Color;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartGradientStop {
    pub(crate) inner: Arc<Mutex<xlsx::ChartGradientStop>>,
}

#[wasm_bindgen]
impl ChartGradientStop {
    #[wasm_bindgen(constructor)]
    pub fn new(color: Color, position: u8) -> ChartGradientStop {
        ChartGradientStop {
            inner: Arc::new(Mutex::new(xlsx::ChartGradientStop::new(
                color.inner,
                position,
            ))),
        }
    }
}
