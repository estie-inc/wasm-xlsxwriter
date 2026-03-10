use crate::wrapper::ChartFont;
use crate::wrapper::ChartTrendlineType;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartTrendline {
    pub(crate) inner: Arc<Mutex<xlsx::ChartTrendline>>,
}

#[wasm_bindgen]
impl ChartTrendline {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ChartTrendline {
        ChartTrendline {
            inner: Arc::new(Mutex::new(xlsx::ChartTrendline::new())),
        }
    }
    #[wasm_bindgen(js_name = "setType", skip_jsdoc)]
    pub fn set_type(&self, trend: ChartTrendlineType) -> ChartTrendline {
        let mut lock = self.inner.lock().unwrap();
        lock.set_type(xlsx::ChartTrendlineType::from(trend));
        ChartTrendline {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setLabelFont", skip_jsdoc)]
    pub fn set_label_font(&self, font: ChartFont) -> ChartTrendline {
        let mut lock = self.inner.lock().unwrap();
        lock.set_label_font(&font.inner);
        ChartTrendline {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setName", skip_jsdoc)]
    pub fn set_name(&self, name: &str) -> ChartTrendline {
        let mut lock = self.inner.lock().unwrap();
        lock.set_name(name);
        ChartTrendline {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setForwardPeriod", skip_jsdoc)]
    pub fn set_forward_period(&self, period: f64) -> ChartTrendline {
        let mut lock = self.inner.lock().unwrap();
        lock.set_forward_period(period);
        ChartTrendline {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setBackwardPeriod", skip_jsdoc)]
    pub fn set_backward_period(&self, period: f64) -> ChartTrendline {
        let mut lock = self.inner.lock().unwrap();
        lock.set_backward_period(period);
        ChartTrendline {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "displayEquation", skip_jsdoc)]
    pub fn display_equation(&self, enable: bool) -> ChartTrendline {
        let mut lock = self.inner.lock().unwrap();
        lock.display_equation(enable);
        ChartTrendline {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "displayRSquared", skip_jsdoc)]
    pub fn display_r_squared(&self, enable: bool) -> ChartTrendline {
        let mut lock = self.inner.lock().unwrap();
        lock.display_r_squared(enable);
        ChartTrendline {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setIntercept", skip_jsdoc)]
    pub fn set_intercept(&self, intercept: f64) -> ChartTrendline {
        let mut lock = self.inner.lock().unwrap();
        lock.set_intercept(intercept);
        ChartTrendline {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "deleteFromLegend", skip_jsdoc)]
    pub fn delete_from_legend(&self, enable: bool) -> ChartTrendline {
        let mut lock = self.inner.lock().unwrap();
        lock.delete_from_legend(enable);
        ChartTrendline {
            inner: Arc::clone(&self.inner),
        }
    }
}
