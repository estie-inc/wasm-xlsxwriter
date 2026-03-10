use crate::wrapper::ChartLayout;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartPlotArea {
    pub(crate) parent: Arc<Mutex<xlsx::Chart>>,
}

#[wasm_bindgen]
impl ChartPlotArea {
    #[wasm_bindgen(js_name = "setLayout", skip_jsdoc)]
    pub fn set_layout(&self, layout: ChartLayout) -> ChartPlotArea {
        let mut lock = self.parent.lock().unwrap();
        lock.plot_area().set_layout(&layout.inner);
        ChartPlotArea {
            parent: Arc::clone(&self.parent),
        }
    }
}
