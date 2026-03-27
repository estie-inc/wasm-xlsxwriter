use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ChartMarkerType` enum defines the {@link Chart} marker types.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ChartMarkerType {
    /// Square marker type.
    Square,
    /// Diamond marker type.
    Diamond,
    /// Triangle marker type.
    Triangle,
    /// X shape marker type.
    X,
    /// Star (X overlaid on vertical dash) marker type.
    Star,
    /// Short dash marker type. This is also called `dot` in some Excel versions.
    ShortDash,
    /// Long dash marker type. This is also called `dash` in some Excel versions.
    LongDash,
    /// Circle marker type.
    Circle,
    /// Plus sign marker type.
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
