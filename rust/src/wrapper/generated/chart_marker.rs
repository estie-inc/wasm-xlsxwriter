use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartMarker {
    pub(crate) inner: Arc<Mutex<xlsx::ChartMarker>>,
}

#[wasm_bindgen]
impl ChartMarker {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ChartMarker {
        ChartMarker {
            inner: Arc::new(Mutex::new(xlsx::ChartMarker::new())),
        }
    }
    #[wasm_bindgen(js_name = "setAutomatic", skip_jsdoc)]
    pub fn set_automatic(&self) -> ChartMarker {
        let mut lock = self.inner.lock().unwrap();
        lock.set_automatic();
        ChartMarker {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setNone", skip_jsdoc)]
    pub fn set_none(&self) -> ChartMarker {
        let mut lock = self.inner.lock().unwrap();
        lock.set_none();
        ChartMarker {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setSize", skip_jsdoc)]
    pub fn set_size(&self, size: u8) -> ChartMarker {
        let mut lock = self.inner.lock().unwrap();
        lock.set_size(size);
        ChartMarker {
            inner: Arc::clone(&self.inner),
        }
    }
}
