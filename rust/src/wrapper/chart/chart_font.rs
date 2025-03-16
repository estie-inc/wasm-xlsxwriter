use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

use crate::wrapper::color::Color;
use crate::macros::wrap_struct;

/// The `ChartFont` struct represents a chart font.
///
/// The `ChartFont` struct is used to define the font properties for chart elements
/// such as chart titles, axis labels, data labels and other text elements in a chart.
///
/// It is used in conjunction with the {@link Chart} struct.

wrap_struct!(
    ChartFont,
    xlsx::ChartFont,
    set_bold(),
    set_character_set(character_set: u8),
    set_color(color: &Color),
    set_italic(),
    set_name(name: &str),
    set_pitch_family(pitch_family: u8),
    set_right_to_left(enable: bool),
    set_rotation(rotation: i16),
    set_size(size: f64),
    set_strikethrough(),
    set_underline(),
    unset_bold()
);
