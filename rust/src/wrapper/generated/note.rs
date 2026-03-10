use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct Note {
    pub(crate) inner: Arc<Mutex<xlsx::Note>>,
}

#[wasm_bindgen]
impl Note {
    #[wasm_bindgen(constructor)]
    pub fn new(text: &str) -> Note {
        Note {
            inner: Arc::new(Mutex::new(xlsx::Note::new(text))),
        }
    }
    #[wasm_bindgen(js_name = "resetText", skip_jsdoc)]
    pub fn reset_text(&self, text: &str) -> Note {
        let mut lock = self.inner.lock().unwrap();
        lock.reset_text(text);
        Note {
            inner: Arc::clone(&self.inner),
        }
    }
}
