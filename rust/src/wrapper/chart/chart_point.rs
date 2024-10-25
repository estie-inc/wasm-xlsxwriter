use rust_xlsxwriter::{self as xlsx};
use wasm_bindgen::prelude::*;

use crate::wrapper::chart::chart_format::ChartFormat;

/// The `ChartPoint` struct represents a chart point.
///
/// The `ChartPoint` struct represents a "point" in a data series which is the
/// element you get in Excel if you right click on an individual data point or
/// segment and select "Format Data Point".
///
/// <img src="https://rustxlsxwriter.github.io/images/chart_point_dialog.png">
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
///
/// TODO: example omitted
#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartPoint {
    pub(crate) inner: xlsx::ChartPoint,
}

#[wasm_bindgen]
impl ChartPoint {
    /// Create a new `ChartPoint` object to represent a Chart point.
    #[wasm_bindgen(constructor)]
    pub fn new() -> ChartPoint {
        ChartPoint {
            inner: xlsx::ChartPoint::new(),
        }
    }

    pub fn set_format(&self, format: &mut ChartFormat) -> ChartPoint {
        ChartPoint {
            inner: self.clone().inner.set_format(&mut format.inner),
        }
    }
}
