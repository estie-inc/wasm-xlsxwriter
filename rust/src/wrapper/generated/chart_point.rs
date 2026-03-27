use crate::wrapper::ChartFormat;
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
    #[doc = r" Create a deep clone of this object."]
    #[wasm_bindgen(js_name = "clone")]
    pub fn deep_clone(&self) -> ChartPoint {
        ChartPoint {
            inner: Arc::new(Mutex::new(self.inner.lock().unwrap().clone())),
        }
    }
    /// Set the formatting properties for a chart point.
    ///
    /// Set the formatting properties for a chart point via a {@link ChartFormat}
    /// object or a sub struct that implements {@link IntoChartFormat}.
    ///
    /// The formatting that can be applied via a {@link ChartFormat} object are:
    ///
    /// - {@link ChartFormat#setSolidFill}: Set the {@link ChartSolidFill} properties.
    /// - {@link ChartFormat#setPatternFill}: Set the {@link ChartPatternFill} properties.
    /// - {@link ChartFormat#setGradientFill}: Set the {@link ChartGradientFill} properties.
    /// - {@link ChartFormat#setNoFill}: Turn off the fill for the chart object.
    /// - {@link ChartFormat#setLine}: Set the {@link ChartLine} properties.
    /// - {@link ChartFormat#setBorder}: Set the {@link ChartBorder} properties.
    ///   A synonym for {@link ChartLine} depending on context.
    /// - {@link ChartFormat#setNoLine}: Turn off the line for the chart object.
    /// - {@link ChartFormat#setNoBorder}: Turn off the border for the chart object.
    ///
    /// # Parameters
    ///
    /// `format`: A {@link ChartFormat} struct reference or a sub struct that will
    /// convert into a `ChartFormat` instance. See the docs for
    /// {@link IntoChartFormat} for details.
    #[wasm_bindgen(js_name = "setFormat", skip_jsdoc)]
    pub fn set_format(&self, format: &ChartFormat) -> ChartPoint {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_format(&mut *format.inner.lock().unwrap());
        *lock = inner;
        ChartPoint {
            inner: Arc::clone(&self.inner),
        }
    }
}
