use rust_xlsxwriter::{self as xlsx};
use wasm_bindgen::prelude::*;

use crate::wrapper::chart::chart_format::ChartFormat;
use crate::wrapper::chart::chart_marker_type::ChartMarkerType;
use crate::macros::wrap_struct;

/// The `ChartMarker` struct represents a chart marker.
///
/// The `ChartMarker` struct is used to define the properties of a chart marker
/// for a data series.
///
/// It is used in conjunction with the {@link Chart} struct.

wrap_struct!(
    ChartMarker,
    xlsx::ChartMarker,
    set_automatic(),
    set_format(format: &mut ChartFormat),
    set_none(),
    set_size(size: u8),
    set_type(marker_type: ChartMarkerType)
);
