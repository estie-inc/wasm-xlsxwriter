use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ShapeGradientFillType` enum defines the gradient types of a
/// {@link ShapeGradientFill}.
///
/// The four gradient types supported by Excel are:
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ShapeGradientFillType {
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

impl From<ShapeGradientFillType> for xlsx::ShapeGradientFillType {
    fn from(value: ShapeGradientFillType) -> xlsx::ShapeGradientFillType {
        match value {
            ShapeGradientFillType::Linear => xlsx::ShapeGradientFillType::Linear,
            ShapeGradientFillType::Radial => xlsx::ShapeGradientFillType::Radial,
            ShapeGradientFillType::Rectangular => xlsx::ShapeGradientFillType::Rectangular,
            ShapeGradientFillType::Path => xlsx::ShapeGradientFillType::Path,
        }
    }
}
