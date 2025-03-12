use std::sync::{Arc, Mutex};

use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

use super::{chart_font::ChartFont, chart_format::ChartFormat, chart_legend_position::ChartLegendPosition};

#[wasm_bindgen]
pub struct ChartLegend {
    pub(crate) chart: Arc<Mutex<xlsx::Chart>>,
}

#[wasm_bindgen]
impl ChartLegend {
    #[wasm_bindgen(js_name = "setHidden", skip_jsdoc)]
    pub fn set_hidden(&self) -> ChartLegend {
        let mut chart = self.chart.lock().unwrap();
        chart.legend().set_hidden();
        ChartLegend {
            chart: Arc::clone(&self.chart),
        }
    }

    #[wasm_bindgen(js_name = "setPosition", skip_jsdoc)]
    pub fn set_position(&self, position: ChartLegendPosition) -> ChartLegend {
        let mut chart = self.chart.lock().unwrap();
        chart.legend().set_position(position.into());
        ChartLegend {
            chart: Arc::clone(&self.chart),
        }
    }

    #[wasm_bindgen(js_name = "setOverlay", skip_jsdoc)]
    pub fn set_overlay(&self, enable: bool) -> ChartLegend {
        let mut chart = self.chart.lock().unwrap();
        chart.legend().set_overlay(enable);
        ChartLegend {
            chart: Arc::clone(&self.chart),
        }
    }

    #[wasm_bindgen(js_name = "setFormat", skip_jsdoc)]
    pub fn set_format(&self, format: &mut ChartFormat) -> ChartLegend {
        let mut chart = self.chart.lock().unwrap();
        chart.legend().set_format(&mut format.inner);
        ChartLegend {
            chart: Arc::clone(&self.chart),
        }
    }

    #[wasm_bindgen(js_name = "setFont", skip_jsdoc)]
    pub fn set_font(&self, font: &ChartFont) -> ChartLegend {
        let mut chart = self.chart.lock().unwrap();
        chart.legend().set_font(&font.inner);
        ChartLegend {
            chart: Arc::clone(&self.chart),
        }
    }
}
