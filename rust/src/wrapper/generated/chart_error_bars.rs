use crate::wrapper::ChartErrorBarsDirection;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ChartErrorBars` struct represents the error bars for a chart series.
///
/// Error bars on Excel charts allow you to show margins of error for a series
/// based on measures such as Standard Deviation, Standard Error, Fixed values,
/// Percentages or even custom defined ranges.
///
/// src="https://rustxlsxwriter.github.io/images/chart_error_bars_options.png">
///
/// The `ChartErrorBars` struct can be added to a series via the
/// {@link ChartSeries#setYErrorBars} and {@link ChartSeries#setXErrorBars}
/// methods.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartErrorBars {
    pub(crate) inner: Arc<Mutex<xlsx::ChartErrorBars>>,
}

#[wasm_bindgen]
impl ChartErrorBars {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ChartErrorBars {
        ChartErrorBars {
            inner: Arc::new(Mutex::new(xlsx::ChartErrorBars::new())),
        }
    }
    /// Set the direction of a Chart series error bars.
    ///
    /// The {@link ChartErrorBarsDirection} enum defines the error bar direction for a
    /// chart series.
    ///
    /// # Parameters
    ///
    /// - `direction`: A {@link ChartErrorBarsDirection} enum reference.
    #[wasm_bindgen(js_name = "setDirection", skip_jsdoc)]
    pub fn set_direction(&self, direction: ChartErrorBarsDirection) -> ChartErrorBars {
        let mut lock = self.inner.lock().unwrap();
        lock.set_direction(xlsx::ChartErrorBarsDirection::from(direction));
        ChartErrorBars {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the end cap on/off for a Chart series error bars.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is on by default.
    #[wasm_bindgen(js_name = "setEndCap", skip_jsdoc)]
    pub fn set_end_cap(&self, enable: bool) -> ChartErrorBars {
        let mut lock = self.inner.lock().unwrap();
        lock.set_end_cap(enable);
        ChartErrorBars {
            inner: Arc::clone(&self.inner),
        }
    }
}
