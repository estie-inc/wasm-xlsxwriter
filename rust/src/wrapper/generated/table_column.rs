use crate::wrapper::Format;
use crate::wrapper::Formula;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct TableColumn {
    pub(crate) inner: Arc<Mutex<xlsx::TableColumn>>,
}

#[wasm_bindgen]
impl TableColumn {
    #[wasm_bindgen(constructor)]
    pub fn new() -> TableColumn {
        TableColumn {
            inner: Arc::new(Mutex::new(xlsx::TableColumn::new())),
        }
    }
    #[wasm_bindgen(js_name = "setHeader", skip_jsdoc)]
    pub fn set_header(&self, caption: &str) -> TableColumn {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_header(caption);
        *lock = inner;
        TableColumn {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setTotalLabel", skip_jsdoc)]
    pub fn set_total_label(&self, label: &str) -> TableColumn {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_total_label(label);
        *lock = inner;
        TableColumn {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setFormula", skip_jsdoc)]
    pub fn set_formula(&self, formula: Formula) -> TableColumn {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_formula(formula.inner.clone());
        *lock = inner;
        TableColumn {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setFormat", skip_jsdoc)]
    pub fn set_format(&self, format: Format) -> TableColumn {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_format(format.inner.clone());
        *lock = inner;
        TableColumn {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setHeaderFormat", skip_jsdoc)]
    pub fn set_header_format(&self, format: Format) -> TableColumn {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_header_format(format.inner.clone());
        *lock = inner;
        TableColumn {
            inner: Arc::clone(&self.inner),
        }
    }
}
