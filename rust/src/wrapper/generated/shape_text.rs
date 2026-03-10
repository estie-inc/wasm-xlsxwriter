use crate::wrapper::ShapeTextDirection;
use crate::wrapper::ShapeTextHorizontalAlignment;
use crate::wrapper::ShapeTextVerticalAlignment;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct ShapeText {
    pub(crate) inner: Arc<Mutex<xlsx::ShapeText>>,
}

#[wasm_bindgen]
impl ShapeText {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ShapeText {
        ShapeText {
            inner: Arc::new(Mutex::new(xlsx::ShapeText::new())),
        }
    }
    #[wasm_bindgen(js_name = "setHorizontalAlignment", skip_jsdoc)]
    pub fn set_horizontal_alignment(&self, alignment: ShapeTextHorizontalAlignment) -> ShapeText {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_horizontal_alignment(xlsx::ShapeTextHorizontalAlignment::from(alignment));
        *lock = inner;
        ShapeText {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setVerticalAlignment", skip_jsdoc)]
    pub fn set_vertical_alignment(&self, alignment: ShapeTextVerticalAlignment) -> ShapeText {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_vertical_alignment(xlsx::ShapeTextVerticalAlignment::from(alignment));
        *lock = inner;
        ShapeText {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setDirection", skip_jsdoc)]
    pub fn set_direction(&self, direction: ShapeTextDirection) -> ShapeText {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_direction(xlsx::ShapeTextDirection::from(direction));
        *lock = inner;
        ShapeText {
            inner: Arc::clone(&self.inner),
        }
    }
}
