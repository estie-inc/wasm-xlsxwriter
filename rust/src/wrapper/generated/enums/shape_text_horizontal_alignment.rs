use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ShapeTextHorizontalAlignment` enum defines the horizontal alignment for
/// Shape text.
///
/// See ShapeText.setHorizontalAlignment().
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify, Default)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ShapeTextHorizontalAlignment {
    /// Horizontally align text in the default position (usually to the left).
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_text_horizontal_alignment_default.png">
    #[default]
    Default,
    /// Horizontally align text to the left of the shape.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_text_horizontal_alignment_default.png">
    Left,
    /// Horizontally align text to the center of the shape.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_text_horizontal_alignment_center.png">
    Center,
    /// Horizontally align text to the right of the shape.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_text_horizontal_alignment_right.png">
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
