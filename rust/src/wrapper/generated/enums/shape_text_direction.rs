use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ShapeTextDirection` enum defines the text direction for {@link Shape} text.
///
/// See {@link ShapeText#setDirection}.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify, Default)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ShapeTextDirection {
    /// Text is horizontal. This is the Excel default.
    #[default]
    Horizontal,
    /// Text is rotated 270 degrees.
    Rotate270,
    /// Text is rotated 90 degrees.
    Rotate90,
    /// Text direction is rotated 90 degrees but the characters aren't rotated. Suitable for East Asian text.
    Rotate90EastAsian,
    /// Text is stacked vertically.
    Stacked,
}

impl From<ShapeTextDirection> for xlsx::ShapeTextDirection {
    fn from(value: ShapeTextDirection) -> xlsx::ShapeTextDirection {
        match value {
            ShapeTextDirection::Horizontal => xlsx::ShapeTextDirection::Horizontal,
            ShapeTextDirection::Rotate270 => xlsx::ShapeTextDirection::Rotate270,
            ShapeTextDirection::Rotate90 => xlsx::ShapeTextDirection::Rotate90,
            ShapeTextDirection::Rotate90EastAsian => xlsx::ShapeTextDirection::Rotate90EastAsian,
            ShapeTextDirection::Stacked => xlsx::ShapeTextDirection::Stacked,
        }
    }
}
