use rust_xlsxwriter::{self as xlsx};
use wasm_bindgen::prelude::*;
use crate::wrapper::chart::chart_line::ChartLine;
use crate::wrapper::chart::chart_solid_fill::ChartSolidFill;

/// The `ChartFormat` struct represents formatting for various chart objects.
///
/// Excel uses a standard formatting dialog for the elements of a chart such as
/// data series, the plot area, the chart area, the legend or individual points.
/// It looks like this:
///
/// <img src="https://rustxlsxwriter.github.io/images/chart_format_dialog.png">
///
/// The {@link ChartFormat} struct represents many of these format options and just
/// like Excel it offers a similar formatting interface for a number of the
/// chart sub-elements supported by `rust_xlsxwriter`.
///
/// It is used in conjunction with the {@link Chart} struct.
///
/// The {@link ChartFormat} struct is generally passed to the `set_format()` method
/// of a chart element. It supports several chart formatting elements such as:
///
/// - {@link ChartFormat#setSolidFill}: Set the {@link ChartSolidFill} properties.
/// - {@link ChartFormat#setPatternFill}: Set the {@link ChartPatternFill}
///   properties.
/// - {@link ChartFormat#setGradientFill}: Set the {@link ChartGradientFill}
///   properties.
/// - {@link ChartFormat#setNoFill}: Turn off the fill for the chart object.
/// - {@link ChartFormat#setLine}: Set the {@link ChartLine} properties.
/// - {@link ChartFormat#setBorder}: Set the {@link ChartBorder} properties. A
///   synonym for {@link ChartLine} depending on context.
/// - {@link ChartFormat#setNoLine}: Turn off the line for the chart object.
/// - {@link ChartFormat#setNoBorder}: Turn off the border for the chart
///   object.
///
/// TODO: example omitted
#[wasm_bindgen]
#[derive(Clone)]
pub struct ChartFormat {
    pub(crate) inner: xlsx::ChartFormat,
}

#[wasm_bindgen]
impl ChartFormat {
    /// Create a new `ChartFormat` instance to set formatting for a chart element.
    #[wasm_bindgen(constructor)]
    pub fn new() -> ChartFormat {
        ChartFormat {
            inner: xlsx::ChartFormat::new(),
        }
    }

    #[wasm_bindgen(js_name = "setLine")]
    pub fn set_line(&mut self, line: &ChartLine) -> ChartFormat {
        self.inner.set_line(&line.inner);
        self.clone()
    }

    #[wasm_bindgen(js_name = "setBorder")]
    pub fn set_border(&mut self, border: &ChartLine) -> ChartFormat {
        self.inner.set_border(&border.inner);
        self.clone()
    }

    #[wasm_bindgen(js_name = "setNoLine")]
    pub fn set_no_line(&mut self) -> ChartFormat {
        self.inner.set_no_line();
        self.clone()
    }

    #[wasm_bindgen(js_name = "setNoBorder")]
    pub fn set_no_border(&mut self) -> ChartFormat {
        self.inner.set_no_border();
        self.clone()
    }

    #[wasm_bindgen(js_name = "setNoFill")]
    pub fn set_no_fill(&mut self) -> ChartFormat {
        self.inner.set_no_fill();
        self.clone()
    }

    // TODO: set_gradient_fill, set_pattern_fill
    #[wasm_bindgen(js_name = "setSolidFill")]
    pub fn set_solid_fill(&mut self, fill: &ChartSolidFill) -> ChartFormat {
        self.inner.set_solid_fill(&fill.inner);
        self.clone()
    }
}
