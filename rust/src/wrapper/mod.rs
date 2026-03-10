mod chart;
mod color;
mod datetime;
mod excel_data;
mod format;
pub mod generated;
mod image;
mod note;
mod rich_string;
mod table;
mod utils;
mod workbook;
mod worksheet;

use crate::error::XlsxError;
use wasm_bindgen::prelude::wasm_bindgen;


pub(crate) type WasmResult<T> = std::result::Result<T, XlsxError>;

// Re-export all types so generated code can use `crate::wrapper::TypeName`
pub(crate) use chart::{
    chart_axis::ChartAxis,
    chart_data_label::ChartDataLabel,
    chart_data_label_position::ChartDataLabelPosition,
    chart_empty_cells::ChartEmptyCells,
    chart_font::ChartFont,
    chart_format::ChartFormat,
    chart_gradient_fill::ChartGradientFill,
    chart_gradient_fill_type::ChartGradientFillType,
    chart_gradient_stop::ChartGradientStop,
    chart_layout::ChartLayout,
    chart_legend::ChartLegend,
    chart_legend_position::ChartLegendPosition,
    chart_line::{ChartLine, ChartLineDashType},
    chart_marker::ChartMarker,
    chart_marker_type::ChartMarkerType,
    chart_pattern_fill::ChartPatternFill,
    chart_pattern_fill_type::ChartPatternFillType,
    chart_point::ChartPoint,
    chart_range::ChartRange,
    chart_series::ChartSeries,
    chart_solid_fill::ChartSolidFill,
    chart_title::ChartTitle,
    chart_type::ChartType,
    Chart,
};
pub(crate) use color::Color;
pub(crate) use datetime::ExcelDateTime;
pub(crate) use format::{
    FontScheme, Format, FormatAlign, FormatBorder, FormatDiagonalBorder, FormatPattern,
    FormatScript, FormatUnderline,
};
pub(crate) use generated::*;
pub(crate) use image::Image;
pub(crate) use note::Note;
pub(crate) use table::{Table, TableColumn, TableStyle};
pub(crate) use workbook::Workbook;
pub(crate) use worksheet::Worksheet;

// This runs once when the wasm module is instantiated
// https://rustwasm.github.io/wasm-bindgen/reference/attributes/on-rust-exports/start.html
#[wasm_bindgen(start)]
pub fn start() {
    console_error_panic_hook::set_once();
}
