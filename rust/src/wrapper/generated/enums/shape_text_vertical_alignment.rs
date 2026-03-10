use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ShapeTextVerticalAlignment` enum defines the vertical alignment for
/// Shape text.
///
/// See ShapeText.setHorizontalAlignment().
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify, Default)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ShapeTextVerticalAlignment {
    /// Vertically align text to the top of the shape. This is the default.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_text_horizontal_alignment_default.png">
    #[default]
    Top,
    /// Vertically align text to the middle of the shape.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_text_vertical_alignment_middle.png">
    Middle,
    /// Vertically align text to the bottom of the shape.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_text_vertical_alignment_bottom.png">
    Bottom,
    /// Vertically align text to the top center of the shape.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_text_vertical_alignment_top_centered.png">
    TopCentered,
    /// Vertically align text to the middle center of the shape.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_text_vertical_alignment_middle_centered.png">
    MiddleCentered,
    /// Vertically align text to the bottom center of the shape.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_text_vertical_alignment_bottom_centered.png">
    BottomCentered,
}

impl From<ShapeTextVerticalAlignment> for xlsx::ShapeTextVerticalAlignment {
    fn from(value: ShapeTextVerticalAlignment) -> xlsx::ShapeTextVerticalAlignment {
        match value {
            ShapeTextVerticalAlignment::Top => xlsx::ShapeTextVerticalAlignment::Top,
            ShapeTextVerticalAlignment::Middle => xlsx::ShapeTextVerticalAlignment::Middle,
            ShapeTextVerticalAlignment::Bottom => xlsx::ShapeTextVerticalAlignment::Bottom,
            ShapeTextVerticalAlignment::TopCentered => {
                xlsx::ShapeTextVerticalAlignment::TopCentered
            }
            ShapeTextVerticalAlignment::MiddleCentered => {
                xlsx::ShapeTextVerticalAlignment::MiddleCentered
            }
            ShapeTextVerticalAlignment::BottomCentered => {
                xlsx::ShapeTextVerticalAlignment::BottomCentered
            }
        }
    }
}
