use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct Button {
    pub(crate) inner: Arc<Mutex<xlsx::Button>>,
}

#[wasm_bindgen]
impl Button {
    #[wasm_bindgen(constructor)]
    pub fn new() -> Button {
        Button {
            inner: Arc::new(Mutex::new(xlsx::Button::new())),
        }
    }
    #[wasm_bindgen(js_name = "setCaption", skip_jsdoc)]
    pub fn set_caption(&self, caption: &str) -> Button {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_caption(caption);
        *lock = inner;
        Button {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setMacro", skip_jsdoc)]
    pub fn set_macro(&self, name: &str) -> Button {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_macro(name);
        *lock = inner;
        Button {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setWidth", skip_jsdoc)]
    pub fn set_width(&self, width: u32) -> Button {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_width(width);
        *lock = inner;
        Button {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setHeight", skip_jsdoc)]
    pub fn set_height(&self, height: u32) -> Button {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_height(height);
        *lock = inner;
        Button {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setAltText", skip_jsdoc)]
    pub fn set_alt_text(&self, alt_text: &str) -> Button {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_alt_text(alt_text);
        *lock = inner;
        Button {
            inner: Arc::clone(&self.inner),
        }
    }
}
