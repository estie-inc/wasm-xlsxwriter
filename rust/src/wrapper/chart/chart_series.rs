use std::sync::{Arc, Mutex};

use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

use super::{chart_data_label::ChartDataLabel, chart_format::ChartFormat, chart_marker::ChartMarker, chart_point::ChartPoint, chart_range::ChartRange};
use crate::macros::wrap_struct;

/// The `ChartSeries` struct represents a chart series.
///
/// A chart in Excel must contain at least one data series. The `ChartSeries` struct 
/// represents the category and value ranges, and the formatting and options for the chart series.
///
/// It is used in conjunction with the {@link Chart} struct.
///
/// A chart series is usually created via the {@link Chart#addSeries} method.
/// However, if required you can create a standalone `ChartSeries` object and add it 
/// to a chart via the {@link Chart#pushSeries} method.

wrap_struct!(
    ChartSeries,
    xlsx::ChartSeries,
    set_categories(range: &ChartRange),
    set_name(name: &str),
    set_values(range: &ChartRange),
    set_points(points: Vec<ChartPoint>),
    set_data_label(data_label: &ChartDataLabel),
    set_marker(marker: &ChartMarker),
    set_format(format: &mut ChartFormat)
);
