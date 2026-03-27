use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ChartGradientFillType` enum defines the gradient types of a
/// {@link ChartGradientFill}.
///
/// The four gradient types supported by Excel are:
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ChartGradientFillType {
    /// The gradient runs linearly from the top of the area vertically to the
    /// bottom. This is the default.
    Linear,
    /// The gradient runs radially from the bottom right of the area vertically
    /// to the top left.
    Radial,
    /// The gradient runs in a rectangular pattern from the bottom right of the
    /// area vertically to the top left.
    Rectangular,
    /// The gradient runs in a rectangular pattern from the center of the area
    /// to the outer vertices.
    Path,
}

impl From<ChartGradientFillType> for xlsx::ChartGradientFillType {
    fn from(value: ChartGradientFillType) -> xlsx::ChartGradientFillType {
        match value {
            ChartGradientFillType::Linear => xlsx::ChartGradientFillType::Linear,
            ChartGradientFillType::Radial => xlsx::ChartGradientFillType::Radial,
            ChartGradientFillType::Rectangular => xlsx::ChartGradientFillType::Rectangular,
            ChartGradientFillType::Path => xlsx::ChartGradientFillType::Path,
        }
    }
}
