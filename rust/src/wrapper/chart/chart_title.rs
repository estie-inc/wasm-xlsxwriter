use std::sync::{Arc, Mutex};

use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

use crate::wrapper::chart::chart_format::ChartFormat;

/// The `ChartTitle` struct represents a chart title.
///
/// It is used in conjunction with the {@link Chart} struct.
#[wasm_bindgen]
pub struct ChartTitle {
    pub(crate) chart: Arc<Mutex<xlsx::Chart>>,
}

#[wasm_bindgen]
impl ChartTitle {
    /// Add a title for a chart.
    ///
    /// Set the name (title) for the chart. The name is displayed above the
    /// chart.
    ///
    /// The name can be a simple string, a formula such as `Sheet1!$A$1` or a
    /// tuple with a sheet name, row and column such as `('Sheet1', 0, 0)`.
    ///
    /// # Parameters
    ///
    /// - `range`: The range property which can be one of the following generic
    ///   types:
    ///    - A simple string title.
    ///    - A string with an Excel like range formula such as `"Sheet1!$A$1"`.
    ///    - A tuple that can be used to create the range programmatically using
    ///      a sheet name and zero indexed row and column values like:
    ///      `("Sheet1", 0, 0)` (this gives the same range as the previous
    ///      string value).
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "setName", skip_jsdoc)]
    pub fn set_name(&self, name: &str) -> ChartTitle {
        let mut chart = self.chart.lock().unwrap();
        chart.title().set_name(name);
        ChartTitle {
            chart: Arc::clone(&self.chart),
        }
    }

    /// Set the formatting properties for a chart title.
    ///
    /// Set the formatting properties for a chart title via a {@link ChartFormat}
    /// object or a sub struct that implements [`IntoChartFormat`].
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
    /// [`IntoChartFormat`] for details.
    ///
    #[wasm_bindgen(js_name = "setFormat")]
    pub fn set_format(&self, format: &mut ChartFormat) -> ChartTitle {
        let mut chart = self.chart.lock().unwrap();
        chart.title().set_format(&mut format.inner);
        ChartTitle {
            chart: Arc::clone(&self.chart),
        }
    }
}
