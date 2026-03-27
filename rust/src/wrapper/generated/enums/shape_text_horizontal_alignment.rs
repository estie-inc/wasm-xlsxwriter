use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ShapeTextHorizontalAlignment` enum defines the horizontal alignment for
/// {@link Shape} text.
///
/// See {@link ShapeText#setHorizontalAlignment}.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify, Default)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ShapeTextHorizontalAlignment {
    /// Horizontally align text in the default position (usually to the left).
    #[default]
    Default,
    /// Horizontally align text to the left of the shape.
    Left,
    /// Horizontally align text to the center of the shape.
    Center,
    /// Horizontally align text to the right of the shape.
    Right,
}

impl From<ShapeTextHorizontalAlignment> for xlsx::ShapeTextHorizontalAlignment {
    fn from(value: ShapeTextHorizontalAlignment) -> xlsx::ShapeTextHorizontalAlignment {
        match value {
            ShapeTextHorizontalAlignment::Default => xlsx::ShapeTextHorizontalAlignment::Default,
            ShapeTextHorizontalAlignment::Left => xlsx::ShapeTextHorizontalAlignment::Left,
            ShapeTextHorizontalAlignment::Center => xlsx::ShapeTextHorizontalAlignment::Center,
            ShapeTextHorizontalAlignment::Right => xlsx::ShapeTextHorizontalAlignment::Right,
        }
    }
}
