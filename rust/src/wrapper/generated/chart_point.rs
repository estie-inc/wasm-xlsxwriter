use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ChartPoint` struct represents a chart point.
///
/// The {@link ChartPoint} struct represents a "point" in a data series which is the
/// element you get in Excel if you right click on an individual data point or
/// segment and select "Format Data Point".
///
/// The meaning of "point" varies between chart types. For a Line chart a point
/// is a line segment; in a Column chart a point is a an individual bar; and in
/// a Pie chart a point is a pie segment.
///
/// Chart points are most commonly used for Pie and Doughnut charts to format
/// individual segments of the chart. In all other chart types the formatting
/// happens at the chart series level.
///
/// It is used in conjunction with the {@link Chart} struct.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartPoint {
    pub(crate) inner: Arc<Mutex<xlsx::ChartPoint>>,
}

#[wasm_bindgen]
impl ChartPoint {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ChartPoint {
        ChartPoint {
            inner: Arc::new(Mutex::new(xlsx::ChartPoint::new())),
        }
    }
}
