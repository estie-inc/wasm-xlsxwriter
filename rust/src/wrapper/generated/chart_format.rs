use crate::wrapper::ChartGradientFill;
use crate::wrapper::ChartLine;
use crate::wrapper::ChartPatternFill;
use crate::wrapper::ChartSolidFill;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartFormat {
    pub(crate) inner: Arc<Mutex<xlsx::ChartFormat>>,
}

#[wasm_bindgen]
impl ChartFormat {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ChartFormat {
        ChartFormat {
            inner: Arc::new(Mutex::new(xlsx::ChartFormat::new())),
        }
    }
    #[wasm_bindgen(js_name = "setLine", skip_jsdoc)]
    pub fn set_line(&self, line: ChartLine) -> ChartFormat {
        let mut lock = self.inner.lock().unwrap();
        lock.set_line(&line.inner);
        ChartFormat {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setBorder", skip_jsdoc)]
    pub fn set_border(&self, line: ChartLine) -> ChartFormat {
        let mut lock = self.inner.lock().unwrap();
        lock.set_border(&line.inner);
        ChartFormat {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setNoLine", skip_jsdoc)]
    pub fn set_no_line(&self) -> ChartFormat {
        let mut lock = self.inner.lock().unwrap();
        lock.set_no_line();
        ChartFormat {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setNoBorder", skip_jsdoc)]
    pub fn set_no_border(&self) -> ChartFormat {
        let mut lock = self.inner.lock().unwrap();
        lock.set_no_border();
        ChartFormat {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setNoFill", skip_jsdoc)]
    pub fn set_no_fill(&self) -> ChartFormat {
        let mut lock = self.inner.lock().unwrap();
        lock.set_no_fill();
        ChartFormat {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setSolidFill", skip_jsdoc)]
    pub fn set_solid_fill(&self, fill: ChartSolidFill) -> ChartFormat {
        let mut lock = self.inner.lock().unwrap();
        lock.set_solid_fill(&fill.inner);
        ChartFormat {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setPatternFill", skip_jsdoc)]
    pub fn set_pattern_fill(&self, fill: ChartPatternFill) -> ChartFormat {
        let mut lock = self.inner.lock().unwrap();
        lock.set_pattern_fill(&fill.inner);
        ChartFormat {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setGradientFill", skip_jsdoc)]
    pub fn set_gradient_fill(&self, fill: ChartGradientFill) -> ChartFormat {
        let mut lock = self.inner.lock().unwrap();
        lock.set_gradient_fill(&fill.inner);
        ChartFormat {
            inner: Arc::clone(&self.inner),
        }
    }
}
