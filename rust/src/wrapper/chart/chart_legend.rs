use std::sync::{Arc, Mutex};

use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

use super::{chart_format::ChartFormat, chart_legend_position::ChartLegendPosition};

#[wasm_bindgen]
pub struct ChartLegend {
    pub(crate) inner: Arc<Mutex<xlsx::ChartLegend>>,
}

#[wasm_bindgen]
impl ChartLegend {
    #[wasm_bindgen(js_name = "setHidden", skip_jsdoc)]
    pub fn set_hidden(&self) -> ChartLegend {
        let mut legend = self.inner.lock().unwrap();
        legend.set_hidden();
        ChartLegend {
            inner: Arc::clone(&self.inner),
        }
    }

    #[wasm_bindgen(js_name = "setPosition", skip_jsdoc)]
    pub fn set_position(&self, position: ChartLegendPosition) -> ChartLegend {
        let mut legend = self.inner.lock().unwrap();
        legend.set_position(position.into());
        ChartLegend {
            inner: Arc::clone(&self.inner),
        }
    }

    #[wasm_bindgen(js_name = "setOverlay", skip_jsdoc)]
    pub fn set_overlay(&self, enable: bool) -> ChartLegend {
        let mut legend = self.inner.lock().unwrap();
        legend.set_overlay(enable);
        ChartLegend {
            inner: Arc::clone(&self.inner),
        }
    }

    #[wasm_bindgen(js_name = "setFormat", skip_jsdoc)]
    pub fn set_format(&self, format: &mut ChartFormat) -> ChartLegend {
        let mut legend = self.inner.lock().unwrap();
        legend.set_format(&mut format.inner);
        ChartLegend {
            inner: Arc::clone(&self.inner),
        }
    }
}
