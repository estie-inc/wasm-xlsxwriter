use rust_xlsxwriter::{self as xlsx};
use wasm_bindgen::prelude::*;

use crate::wrapper::chart::chart_data_label_position::ChartDataLabelPosition;
use crate::wrapper::chart::chart_font::ChartFont;
use crate::wrapper::chart::chart_format::ChartFormat;
use crate::macros::wrap_struct;

wrap_struct!(
    ChartDataLabel,
    xlsx::ChartDataLabel,
    show_value(),
    show_category_name(),
    show_series_name(),
    show_leader_lines(),
    show_legend_key(),
    show_percentage(),
    set_position(position: ChartDataLabelPosition),
    set_font(font: &ChartFont),
    set_format(format: &mut ChartFormat),
    set_num_format(num_format: &str),
    set_separator(separator: char),
    show_y_value(),
    show_x_value(),
    set_hidden(),
    set_value(value: &str),
    to_custom()
);

