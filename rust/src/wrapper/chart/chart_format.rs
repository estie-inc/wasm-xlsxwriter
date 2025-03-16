use rust_xlsxwriter::{self as xlsx};
use wasm_bindgen::prelude::*;

use crate::wrapper::color::Color;
use crate::wrapper::chart::chart_solid_fill::ChartSolidFill;
use crate::macros::wrap_struct;

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

wrap_struct!(
    ChartFormat,
    xlsx::ChartFormat,
    set_line(line: &ChartLine),
    set_border(border: &ChartLine),
    set_no_line(),
    set_no_border(),
    set_no_fill(),
    set_solid_fill(fill: &ChartSolidFill)
);

wrap_struct!(
    ChartLine,
    xlsx::ChartLine,
    set_color(color: &Color),
    set_width(width: f64),
    set_dash_type(dash_type: ChartLineDashType),
    set_transparency(transparency: u8),
    set_hidden(enable: bool)
);

/// The `ChartLineDashType` enum defines the dash types for chart lines and borders.
///
/// The `ChartLineDashType` enum defines the dash types for chart lines and
/// borders. It is used in conjunction with the {@link ChartLine#setDashType}
/// method.
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
