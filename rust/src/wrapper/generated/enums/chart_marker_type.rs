use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ChartMarkerType` enum defines the Chart marker types.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ChartMarkerType {
    /// Square marker type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_marker_type_square.png">
    Square,
    /// Diamond marker type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_marker_type_diamond.png">
    Diamond,
    /// Triangle marker type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_marker_type_triangle.png">
    Triangle,
    /// X shape marker type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_marker_type_x.png">
    X,
    /// Star (X overlaid on vertical dash) marker type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_marker_type_star.png">
    Star,
    /// Short dash marker type. This is also called `dot` in some Excel versions.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_marker_type_short_dash.png">
    ShortDash,
    /// Long dash marker type. This is also called `dash` in some Excel versions.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_marker_type_long_dash.png">
    LongDash,
    /// Circle marker type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_marker_type_circle.png">
    Circle,
    /// Plus sign marker type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_marker_type_plus_sign.png">
    PlusSign,
}

impl From<ChartMarkerType> for xlsx::ChartMarkerType {
    fn from(value: ChartMarkerType) -> xlsx::ChartMarkerType {
        match value {
            ChartMarkerType::Square => xlsx::ChartMarkerType::Square,
            ChartMarkerType::Diamond => xlsx::ChartMarkerType::Diamond,
            ChartMarkerType::Triangle => xlsx::ChartMarkerType::Triangle,
            ChartMarkerType::X => xlsx::ChartMarkerType::X,
            ChartMarkerType::Star => xlsx::ChartMarkerType::Star,
            ChartMarkerType::ShortDash => xlsx::ChartMarkerType::ShortDash,
            ChartMarkerType::LongDash => xlsx::ChartMarkerType::LongDash,
            ChartMarkerType::Circle => xlsx::ChartMarkerType::Circle,
            ChartMarkerType::PlusSign => xlsx::ChartMarkerType::PlusSign,
        }
    }
}
