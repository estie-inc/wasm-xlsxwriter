use crate::wrapper::ChartLayout;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ChartPlotArea` struct is a representation of the plotting area an Excel
/// chart.
///
/// The `ChartPlotArea` struct can be used to configure properties of the chart
/// plot area such as the formatting and layout and is usually obtained via the
/// {@link Chart#plotArea} method.
///
/// It is used in conjunction with the {@link Chart} struct.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartPlotArea {
    pub(crate) parent: Arc<Mutex<xlsx::Chart>>,
}

#[wasm_bindgen]
impl ChartPlotArea {
    /// Set the manual position of the plot area.
    ///
    /// This method is used to simulate manual positioning of a chart plot area.
    /// See {@link ChartLayout} for more details.
    ///
    /// # Parameters
    ///
    /// - `layout`: A {@link ChartLayout} struct reference.
    #[wasm_bindgen(js_name = "setLayout", skip_jsdoc)]
    pub fn set_layout(&self, layout: ChartLayout) -> ChartPlotArea {
        let mut lock = self.parent.lock().unwrap();
        lock.plot_area().set_layout(&layout.inner);
        ChartPlotArea {
            parent: Arc::clone(&self.parent),
        }
    }
}
