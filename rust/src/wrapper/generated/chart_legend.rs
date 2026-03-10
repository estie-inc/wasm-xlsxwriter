use crate::wrapper::ChartFont;
use crate::wrapper::ChartLayout;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartLegend {
    pub(crate) parent: Arc<Mutex<xlsx::Chart>>,
}

#[wasm_bindgen]
impl ChartLegend {
    #[wasm_bindgen(js_name = "setHidden", skip_jsdoc)]
    pub fn set_hidden(&self) -> ChartLegend {
        let mut lock = self.parent.lock().unwrap();
        lock.legend().set_hidden();
        ChartLegend {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setOverlay", skip_jsdoc)]
    pub fn set_overlay(&self, enable: bool) -> ChartLegend {
        let mut lock = self.parent.lock().unwrap();
        lock.legend().set_overlay(enable);
        ChartLegend {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setLayout", skip_jsdoc)]
    pub fn set_layout(&self, layout: ChartLayout) -> ChartLegend {
        let mut lock = self.parent.lock().unwrap();
        lock.legend().set_layout(&layout.inner);
        ChartLegend {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setFont", skip_jsdoc)]
    pub fn set_font(&self, font: ChartFont) -> ChartLegend {
        let mut lock = self.parent.lock().unwrap();
        lock.legend().set_font(&font.inner);
        ChartLegend {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "deleteEntries", skip_jsdoc)]
    pub fn delete_entries(&self, entries: Vec<usize>) -> ChartLegend {
        let mut lock = self.parent.lock().unwrap();
        lock.legend().delete_entries(&entries);
        ChartLegend {
            parent: Arc::clone(&self.parent),
        }
    }
}
