use crate::wrapper::ChartFormat;
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
    /// Set the formatting properties for the plot area.
    ///
    /// Set the formatting properties for a chart plot area via a
    /// {@link ChartFormat} object. In Excel the plot area is the area between the
    /// axes on which the chart series are plotted.
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
    /// - `format`: A {@link ChartFormat} struct reference or a sub struct that will
    ///   convert into a `ChartFormat` instance. See the docs for
    ///   {@link IntoChartFormat} for details.
    #[wasm_bindgen(js_name = "setFormat", skip_jsdoc)]
    pub fn set_format(&self, format: &ChartFormat) -> ChartPlotArea {
        let mut lock = self.parent.lock().unwrap();
        lock.plot_area()
            .set_format(&mut *format.inner.lock().unwrap());
        ChartPlotArea {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the manual position of the plot area.
    ///
    /// This method is used to simulate manual positioning of a chart plot area.
    /// See {@link ChartLayout} for more details.
    ///
    /// # Parameters
    ///
    /// - `layout`: A {@link ChartLayout} struct reference.
    #[wasm_bindgen(js_name = "setLayout", skip_jsdoc)]
    pub fn set_layout(&self, layout: &ChartLayout) -> ChartPlotArea {
        let mut lock = self.parent.lock().unwrap();
        lock.plot_area().set_layout(&*layout.inner.lock().unwrap());
        ChartPlotArea {
            parent: Arc::clone(&self.parent),
        }
    }
}
