use std::sync::Arc;

use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;
use wasm_bindgen::prelude::*;

use crate::wrapper::ChartErrorBars;

/// Subset of ChartErrorBarsType without the Custom variant
/// (Custom requires ChartRange struct refs which can't be expressed as a tsify enum).
#[derive(Debug, Clone, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ChartErrorBarsType {
    FixedValue(f64),
    Percentage(f64),
    StandardDeviation(f64),
    StandardError,
}

impl From<ChartErrorBarsType> for xlsx::ChartErrorBarsType {
    fn from(t: ChartErrorBarsType) -> Self {
        match t {
            ChartErrorBarsType::FixedValue(v) => xlsx::ChartErrorBarsType::FixedValue(v),
            ChartErrorBarsType::Percentage(v) => xlsx::ChartErrorBarsType::Percentage(v),
            ChartErrorBarsType::StandardDeviation(v) => {
                xlsx::ChartErrorBarsType::StandardDeviation(v)
            }
            ChartErrorBarsType::StandardError => xlsx::ChartErrorBarsType::StandardError,
        }
    }
}

#[wasm_bindgen]
impl ChartErrorBars {
    #[wasm_bindgen(js_name = "setType")]
    pub fn set_type(&self, error_bars_type: ChartErrorBarsType) -> ChartErrorBars {
        let mut lock = self.inner.lock().unwrap();
        lock.set_type(xlsx::ChartErrorBarsType::from(error_bars_type));
        ChartErrorBars {
            inner: Arc::clone(&self.inner),
        }
    }
}
