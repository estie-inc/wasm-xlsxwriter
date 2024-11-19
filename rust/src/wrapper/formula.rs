use std::sync::{Arc, Mutex};

use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct Formula {
    pub(crate) inner: Arc<Mutex<xlsx::Formula>>,
}

macro_rules! impl_method {
    ($self:ident.$method:ident($($arg:expr),*)) => {
        let mut lock = $self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Formula::new(""));
        inner = inner.$method($($arg),*);
        let _ = std::mem::replace(&mut *lock, inner);
        return Formula {
            inner: Arc::clone(&$self.inner),
        }
    };
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
        impl_method!(self.set_result(result));
    }
}
