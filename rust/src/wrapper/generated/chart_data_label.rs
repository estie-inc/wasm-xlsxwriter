use crate::wrapper::ChartFont;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartDataLabel {
    pub(crate) inner: Arc<Mutex<xlsx::ChartDataLabel>>,
}

#[wasm_bindgen]
impl ChartDataLabel {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ChartDataLabel {
        ChartDataLabel {
            inner: Arc::new(Mutex::new(xlsx::ChartDataLabel::new())),
        }
    }
    #[wasm_bindgen(js_name = "showValue", skip_jsdoc)]
    pub fn show_value(&self) -> ChartDataLabel {
        let mut lock = self.inner.lock().unwrap();
        lock.show_value();
        ChartDataLabel {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "showCategoryName", skip_jsdoc)]
    pub fn show_category_name(&self) -> ChartDataLabel {
        let mut lock = self.inner.lock().unwrap();
        lock.show_category_name();
        ChartDataLabel {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "showSeriesName", skip_jsdoc)]
    pub fn show_series_name(&self) -> ChartDataLabel {
        let mut lock = self.inner.lock().unwrap();
        lock.show_series_name();
        ChartDataLabel {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "showLeaderLines", skip_jsdoc)]
    pub fn show_leader_lines(&self) -> ChartDataLabel {
        let mut lock = self.inner.lock().unwrap();
        lock.show_leader_lines();
        ChartDataLabel {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "showLegendKey", skip_jsdoc)]
    pub fn show_legend_key(&self) -> ChartDataLabel {
        let mut lock = self.inner.lock().unwrap();
        lock.show_legend_key();
        ChartDataLabel {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "showPercentage", skip_jsdoc)]
    pub fn show_percentage(&self) -> ChartDataLabel {
        let mut lock = self.inner.lock().unwrap();
        lock.show_percentage();
        ChartDataLabel {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setFont", skip_jsdoc)]
    pub fn set_font(&self, font: ChartFont) -> ChartDataLabel {
        let mut lock = self.inner.lock().unwrap();
        lock.set_font(&font.inner);
        ChartDataLabel {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setNumFormat", skip_jsdoc)]
    pub fn set_num_format(&self, num_format: &str) -> ChartDataLabel {
        let mut lock = self.inner.lock().unwrap();
        lock.set_num_format(num_format);
        ChartDataLabel {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "showYValue", skip_jsdoc)]
    pub fn show_y_value(&self) -> ChartDataLabel {
        let mut lock = self.inner.lock().unwrap();
        lock.show_y_value();
        ChartDataLabel {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "showXValue", skip_jsdoc)]
    pub fn show_x_value(&self) -> ChartDataLabel {
        let mut lock = self.inner.lock().unwrap();
        lock.show_x_value();
        ChartDataLabel {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setHidden", skip_jsdoc)]
    pub fn set_hidden(&self) -> ChartDataLabel {
        let mut lock = self.inner.lock().unwrap();
        lock.set_hidden();
        ChartDataLabel {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "toCustom", skip_jsdoc)]
    pub fn to_custom(&self) -> ChartDataLabel {
        let mut lock = self.inner.lock().unwrap();
        lock.to_custom();
        ChartDataLabel {
            inner: Arc::clone(&self.inner),
        }
    }
}
