use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

use crate::wrapper::color::Color;
use crate::macros::wrap_struct;

/// The `ChartSolidFill` struct represents a solid fill for chart elements.
///
/// The `ChartSolidFill` struct is used to define the solid fill properties
/// for chart elements such as plot areas, chart areas, data series, and other
/// fillable elements in a chart.
///
/// It is used in conjunction with the {@link Chart} struct.

wrap_struct!(
    ChartSolidFill,
    xlsx::ChartSolidFill,
    set_color(color: &Color),
    set_transparency(transparency: u8)
);
