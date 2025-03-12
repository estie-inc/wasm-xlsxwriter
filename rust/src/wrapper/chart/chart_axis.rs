use std::sync::{Arc, Mutex};

use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

use crate::wrapper::chart::chart_font::ChartFont;

#[derive(Copy, Clone)]
pub enum AxisType {
    X,
    Y,
    X2,
    Y2,
}

impl ChartAxis {
    fn with_chart<F>(&self, f: F) -> ChartAxis
    where
        F: FnOnce(&mut xlsx::ChartAxis),
    {
        let mut chart = self.inner.lock().unwrap();
        let axis = match self.axis {
            AxisType::X => chart.x_axis(),
            AxisType::Y => chart.y_axis(),
            AxisType::X2 => chart.x2_axis(),
            AxisType::Y2 => chart.y2_axis(),
        };
        f(axis);
        ChartAxis {
            inner: Arc::clone(&self.inner),
            axis: self.axis,
        }
    }
}

#[wasm_bindgen]
pub struct ChartAxis {
    pub(crate) inner: Arc<Mutex<xlsx::Chart>>,
    pub(crate) axis: AxisType,
}

#[wasm_bindgen]
impl ChartAxis {
    #[wasm_bindgen(js_name = "setName", skip_jsdoc)]
    pub fn set_name(&self, name: &str) -> ChartAxis {
        self.with_chart(|axis| {
            axis.set_name(name);
        })
    }

    #[wasm_bindgen(js_name = "setNumFormat", skip_jsdoc)]
    pub fn set_num_format(&self, num_format: &str) -> ChartAxis {
        self.with_chart(|axis| {
            axis.set_num_format(num_format);
        })
    }

    #[wasm_bindgen(js_name = "setMin", skip_jsdoc)]
    pub fn set_min(&self, min: f64) -> ChartAxis {
        self.with_chart(|axis| {
            axis.set_min(min);
        })
    }

    #[wasm_bindgen(js_name = "setMax", skip_jsdoc)]
    pub fn set_max(&self, max: f64) -> ChartAxis {
        self.with_chart(|axis| {
            axis.set_max(max);
        })
    }

    #[wasm_bindgen(js_name = "setFont", skip_jsdoc)]
    pub fn set_font(&self, font: &ChartFont) -> ChartAxis {
        self.with_chart(|axis| {
            axis.set_font(&font.inner);
        })
    }
}
