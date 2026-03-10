use crate::wrapper::ChartFont;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartDataTable {
    pub(crate) inner: Arc<Mutex<xlsx::ChartDataTable>>,
}

#[wasm_bindgen]
impl ChartDataTable {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ChartDataTable {
        ChartDataTable {
            inner: Arc::new(Mutex::new(xlsx::ChartDataTable::new())),
        }
    }
    #[wasm_bindgen(js_name = "showHorizontalBorders", skip_jsdoc)]
    pub fn show_horizontal_borders(&self, enable: bool) -> ChartDataTable {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.show_horizontal_borders(enable);
        *lock = inner;
        ChartDataTable {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "showVerticalBorders", skip_jsdoc)]
    pub fn show_vertical_borders(&self, enable: bool) -> ChartDataTable {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.show_vertical_borders(enable);
        *lock = inner;
        ChartDataTable {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "showOutlineBorders", skip_jsdoc)]
    pub fn show_outline_borders(&self, enable: bool) -> ChartDataTable {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.show_outline_borders(enable);
        *lock = inner;
        ChartDataTable {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "showLegendKeys", skip_jsdoc)]
    pub fn show_legend_keys(&self, enable: bool) -> ChartDataTable {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.show_legend_keys(enable);
        *lock = inner;
        ChartDataTable {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setFont", skip_jsdoc)]
    pub fn set_font(&self, font: ChartFont) -> ChartDataTable {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_font(&font.inner);
        *lock = inner;
        ChartDataTable {
            inner: Arc::clone(&self.inner),
        }
    }
}
