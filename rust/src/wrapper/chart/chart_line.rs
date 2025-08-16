use rust_xlsxwriter::{self as xlsx};
use wasm_bindgen::prelude::wasm_bindgen;
use crate::wrapper::color::Color;

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
#[derive(Clone)]
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
    pub fn set_color(&mut self, color: Color) -> ChartLine {
        self.inner.set_color(color);
        self.clone()
    }

    #[wasm_bindgen(js_name = "setWidth")]
    pub fn set_width(&mut self, width: f64) -> ChartLine {
        self.inner.set_width(width);
        self.clone()
    }

    #[wasm_bindgen(js_name = "setDashType")]
    pub fn set_dash_type(&mut self, dash_type: ChartLineDashType) -> ChartLine {
        self.inner.set_dash_type(dash_type.into());
        self.clone()
    }

    #[wasm_bindgen(js_name = "setTransparency")]
    pub fn set_transparency(&mut self, transparency: u8) -> ChartLine {
        self.inner.set_transparency(transparency);
        self.clone()
    }

    #[wasm_bindgen(js_name = "setHidden")]
    pub fn set_hidden(&mut self, enable: bool) -> ChartLine {
        self.inner.set_hidden(enable);
        self.clone()
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
