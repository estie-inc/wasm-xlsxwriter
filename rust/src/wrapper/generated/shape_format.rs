use crate::wrapper::ShapeGradientFill;
use crate::wrapper::ShapeLine;
use crate::wrapper::ShapePatternFill;
use crate::wrapper::ShapeSolidFill;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct ShapeFormat {
    pub(crate) inner: Arc<Mutex<xlsx::ShapeFormat>>,
}

#[wasm_bindgen]
impl ShapeFormat {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ShapeFormat {
        ShapeFormat {
            inner: Arc::new(Mutex::new(xlsx::ShapeFormat::new())),
        }
    }
    #[wasm_bindgen(js_name = "setLine", skip_jsdoc)]
    pub fn set_line(&self, line: ShapeLine) -> ShapeFormat {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_line(&*line.inner.lock().unwrap());
        *lock = inner;
        ShapeFormat {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setNoLine", skip_jsdoc)]
    pub fn set_no_line(&self) -> ShapeFormat {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_no_line();
        *lock = inner;
        ShapeFormat {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setSolidFill", skip_jsdoc)]
    pub fn set_solid_fill(&self, fill: ShapeSolidFill) -> ShapeFormat {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_solid_fill(&*fill.inner.lock().unwrap());
        *lock = inner;
        ShapeFormat {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setNoFill", skip_jsdoc)]
    pub fn set_no_fill(&self) -> ShapeFormat {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_no_fill();
        *lock = inner;
        ShapeFormat {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setPatternFill", skip_jsdoc)]
    pub fn set_pattern_fill(&self, fill: ShapePatternFill) -> ShapeFormat {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_pattern_fill(&*fill.inner.lock().unwrap());
        *lock = inner;
        ShapeFormat {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setGradientFill", skip_jsdoc)]
    pub fn set_gradient_fill(&self, fill: ShapeGradientFill) -> ShapeFormat {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_gradient_fill(&*fill.inner.lock().unwrap());
        *lock = inner;
        ShapeFormat {
            inner: Arc::clone(&self.inner),
        }
    }
}
