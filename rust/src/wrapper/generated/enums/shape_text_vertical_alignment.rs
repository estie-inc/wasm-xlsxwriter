use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ShapeTextVerticalAlignment` enum defines the vertical alignment for
/// {@link Shape} text.
///
/// See {@link ShapeText#setHorizontalAlignment}.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify, Default)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ShapeTextVerticalAlignment {
    /// Vertically align text to the top of the shape. This is the default.
    #[default]
    Top,
    /// Vertically align text to the middle of the shape.
    Middle,
    /// Vertically align text to the bottom of the shape.
    Bottom,
    /// Vertically align text to the top center of the shape.
    TopCentered,
    /// Vertically align text to the middle center of the shape.
    MiddleCentered,
    /// Vertically align text to the bottom center of the shape.
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
