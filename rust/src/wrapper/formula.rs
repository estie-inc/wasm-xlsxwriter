use std::sync::{Arc, Mutex};

use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct Formula {
    pub(crate) inner: Arc<Mutex<xlsx::Formula>>,
}

#[wasm_bindgen]
impl Formula {
    pub(crate) fn lock(&self) -> std::sync::MutexGuard<'_, xlsx::Formula> {
        self.inner.lock().unwrap()
    }

    #[wasm_bindgen(constructor)]
    pub fn new(formula: &str) -> Formula {
        Formula {
            inner: Arc::new(Mutex::new(xlsx::Formula::new(formula))),
        }
    }

    #[wasm_bindgen(js_name = "setResult")]
    pub fn set_result(&self, result: &str) -> Formula {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Formula::new(""));
        inner = inner.set_result(result);
        let _ = std::mem::replace(&mut *lock, inner);
        Formula {
            inner: Arc::clone(&self.inner),
        }
    }
}
