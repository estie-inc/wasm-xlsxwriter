use rust_xlsxwriter::{self as xlsx};
use wasm_bindgen::prelude::*;

use crate::wrapper::color::Color;
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
    pub fn set_line(mut self, line: &ChartLine) -> ChartFormat {
        self.inner.set_line(&line.inner);
        self
    }

    #[wasm_bindgen(js_name = "setBorder")]
    pub fn set_border(mut self, border: &ChartLine) -> ChartFormat {
        self.inner.set_border(&border.inner);
        self
    }

    #[wasm_bindgen(js_name = "setNoLine")]
    pub fn set_no_line(mut self) -> ChartFormat {
        self.inner.set_no_line();
        self
    }

    #[wasm_bindgen(js_name = "setNoBorder")]
    pub fn set_no_border(mut self) -> ChartFormat {
        self.inner.set_no_border();
        self
    }

    #[wasm_bindgen(js_name = "setNoFill")]
    pub fn set_no_fill(mut self) -> ChartFormat {
        self.inner.set_no_fill();
        self
    }

    // TODO: set_gradient_fill, set_pattern_fill
    #[wasm_bindgen(js_name = "setSolidFill")]
    pub fn set_solid_fill(mut self, fill: &ChartSolidFill) -> ChartFormat {
        self.inner.set_solid_fill(&fill.inner);
        self
    }
}

/// The `ChartLine` struct represents a chart line/border.
///
/// The `ChartLine` struct represents the formatting properties for a line or
/// border for a Chart element. It is a sub property of the {@link ChartFormat}
/// struct and is used with the {@link ChartFormat#setLine} or
/// {@link ChartFormat#setBorder} methods.
///
/// Excel uses the element names "Line" and "Border" depending on the context.
/// For a Line chart the line is represented by a line property but for a Column
/// chart the line becomes the border. Both of these share the same properties
/// and are both represented in `rust_xlsxwriter` by the {@link ChartLine} struct.
///
/// As a syntactic shortcut you can use the type alias {@link ChartBorder} instead
/// of `ChartLine`.
///
/// It is used in conjunction with the {@link Chart} struct.
///
/// TODO: example omitted
#[wasm_bindgen]
pub struct ChartLine {
    pub(crate) inner: xlsx::ChartLine,
}

#[wasm_bindgen]
impl ChartLine {
    /// Create a new `ChartLine` instance to set formatting for a chart line.
    #[wasm_bindgen(constructor)]
    pub fn new() -> ChartLine {
        ChartLine {
            inner: xlsx::ChartLine::new(),
        }
    }

    #[wasm_bindgen(js_name = "setColor")]
    pub fn set_color(mut self, color: &Color) -> ChartLine {
        self.inner.set_color(color.inner);
        self
    }

    #[wasm_bindgen(js_name = "setWidth")]
    pub fn set_width(mut self, width: f64) -> ChartLine {
        self.inner.set_width(width);
        self
    }

    #[wasm_bindgen(js_name = "setDashType")]
    pub fn set_dash_type(mut self, dash_type: ChartLineDashType) -> ChartLine {
        self.inner.set_dash_type(dash_type.into());
        self
    }

    #[wasm_bindgen(js_name = "setTransparency")]
    pub fn set_transparency(mut self, transparency: u8) -> ChartLine {
        self.inner.set_transparency(transparency);
        self
    }

    #[wasm_bindgen(js_name = "setHidden")]
    pub fn set_hidden(mut self, enable: bool) -> ChartLine {
        self.inner.set_hidden(enable);
        self
    }
}

#[derive(Clone, Copy)]
#[wasm_bindgen]
pub enum ChartLineDashType {
    /// Solid - chart line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_solid.png">
    Solid,
    /// Round dot - chart line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_round_dot.png">
    RoundDot,
    /// Square dot - chart line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_square_dot.png">
    SquareDot,
    /// Dash - chart line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_dash.png">
    Dash,
    /// Dash dot - chart line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_dash_dot.png">
    DashDot,
    /// Long dash - chart line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_longdash.png">
    LongDash,
    /// Long dash dot - chart line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_longdash_dot.png">
    LongDashDot,
    /// Long dash dot dot - chart line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_longdash_dot_dot.png">
    LongDashDotDot,
}

impl From<ChartLineDashType> for xlsx::ChartLineDashType {
    fn from(value: ChartLineDashType) -> Self {
        match value {
            ChartLineDashType::Solid => xlsx::ChartLineDashType::Solid,
            ChartLineDashType::RoundDot => xlsx::ChartLineDashType::RoundDot,
            ChartLineDashType::SquareDot => xlsx::ChartLineDashType::SquareDot,
            ChartLineDashType::Dash => xlsx::ChartLineDashType::Dash,
            ChartLineDashType::DashDot => xlsx::ChartLineDashType::DashDot,
            ChartLineDashType::LongDash => xlsx::ChartLineDashType::LongDash,
            ChartLineDashType::LongDashDot => xlsx::ChartLineDashType::LongDashDot,
            ChartLineDashType::LongDashDotDot => xlsx::ChartLineDashType::LongDashDotDot,
        }
    }
}
