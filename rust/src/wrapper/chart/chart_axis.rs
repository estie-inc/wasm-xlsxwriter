use std::sync::{Arc, Mutex};

use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

use super::chart_format::ChartFormat;

#[wasm_bindgen]
pub struct ChartAxis {
    pub(crate) inner: Arc<Mutex<xlsx::ChartAxis>>,
}

#[wasm_bindgen]
impl ChartAxis {
    #[wasm_bindgen(js_name = "setName", skip_jsdoc)]
    pub fn set_name(&self, name: &str) -> ChartAxis {
        let mut axis = self.inner.lock().unwrap();
        axis.set_name(name);
        ChartAxis {
            inner: Arc::clone(&self.inner),
        }
    }

    #[wasm_bindgen(js_name = "setFormat", skip_jsdoc)]
    pub fn set_format(&self, format: &mut ChartFormat) -> ChartAxis {
        let mut axis = self.inner.lock().unwrap();
        axis.set_format(&mut format.inner);
        ChartAxis {
            inner: Arc::clone(&self.inner),
        }
    }

    #[wasm_bindgen(js_name = "setNumFormat", skip_jsdoc)]
    pub fn set_num_format(&self, num_format: &str) -> ChartAxis {
        let mut axis = self.inner.lock().unwrap();
        axis.set_num_format(num_format);
        ChartAxis {
            inner: Arc::clone(&self.inner),
        }
    }

    #[wasm_bindgen(js_name = "setMin", skip_jsdoc)]
    pub fn set_min(&self, min: f64) -> ChartAxis {
        let mut axis = self.inner.lock().unwrap();
        axis.set_min(min);
        ChartAxis {
            inner: Arc::clone(&self.inner),
        }
    }

    #[wasm_bindgen(js_name = "setMax", skip_jsdoc)]
    pub fn set_max(&self, max: f64) -> ChartAxis {
        let mut axis = self.inner.lock().unwrap();
        axis.set_max(max);
        ChartAxis {
            inner: Arc::clone(&self.inner),
        }
    }
}
