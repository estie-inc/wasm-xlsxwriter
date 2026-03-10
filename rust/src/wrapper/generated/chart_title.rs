use crate::wrapper::ChartFont;
use crate::wrapper::ChartLayout;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartTitle {
    pub(crate) parent: Arc<Mutex<xlsx::Chart>>,
}

#[wasm_bindgen]
impl ChartTitle {
    #[wasm_bindgen(js_name = "setHidden", skip_jsdoc)]
    pub fn set_hidden(&self) -> ChartTitle {
        let mut lock = self.parent.lock().unwrap();
        lock.title().set_hidden();
        ChartTitle {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setFont", skip_jsdoc)]
    pub fn set_font(&self, font: ChartFont) -> ChartTitle {
        let mut lock = self.parent.lock().unwrap();
        lock.title().set_font(&font.inner);
        ChartTitle {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setLayout", skip_jsdoc)]
    pub fn set_layout(&self, layout: ChartLayout) -> ChartTitle {
        let mut lock = self.parent.lock().unwrap();
        lock.title().set_layout(&layout.inner);
        ChartTitle {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setOverlay", skip_jsdoc)]
    pub fn set_overlay(&self, enable: bool) -> ChartTitle {
        let mut lock = self.parent.lock().unwrap();
        lock.title().set_overlay(enable);
        ChartTitle {
            parent: Arc::clone(&self.parent),
        }
    }
}
