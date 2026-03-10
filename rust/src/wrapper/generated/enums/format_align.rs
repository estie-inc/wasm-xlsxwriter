use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `FormatAlign` enum defines the vertical and horizontal alignment properties
/// of a {@link Format}.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify, Default)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum FormatAlign {
    /// General/default alignment. The cell will use Excel's default for the
    /// data type, for example Left for text and Right for numbers.
    #[default]
    General,
    /// Align text to the left.
    Left,
    /// Center text horizontally.
    Center,
    /// Align text to the right.
    Right,
    /// Fill (repeat) the text horizontally across the cell.
    Fill,
    /// Aligns the text to the left and right of the cell, if the text exceeds
    /// the width of the cell.
    Justify,
    /// Center the text across the cell or cells that have this alignment. This
    /// is an older form of merged cells.
    CenterAcross,
    /// Distribute the words in the text evenly across the cell.
    Distributed,
    /// Align text to the top.
    Top,
    /// Align text to the bottom.
    Bottom,
    /// Center text vertically.
    VerticalCenter,
    /// Aligns the text to the top and bottom of the cell, if the text exceeds
    /// the height of the cell.
    VerticalJustify,
    /// Distribute the words in the text evenly from top to bottom in the cell.
    VerticalDistributed,
}

impl From<FormatAlign> for xlsx::FormatAlign {
    fn from(value: FormatAlign) -> xlsx::FormatAlign {
        match value {
            FormatAlign::General => xlsx::FormatAlign::General,
            FormatAlign::Left => xlsx::FormatAlign::Left,
            FormatAlign::Center => xlsx::FormatAlign::Center,
            FormatAlign::Right => xlsx::FormatAlign::Right,
            FormatAlign::Fill => xlsx::FormatAlign::Fill,
            FormatAlign::Justify => xlsx::FormatAlign::Justify,
            FormatAlign::CenterAcross => xlsx::FormatAlign::CenterAcross,
            FormatAlign::Distributed => xlsx::FormatAlign::Distributed,
            FormatAlign::Top => xlsx::FormatAlign::Top,
            FormatAlign::Bottom => xlsx::FormatAlign::Bottom,
            FormatAlign::VerticalCenter => xlsx::FormatAlign::VerticalCenter,
            FormatAlign::VerticalJustify => xlsx::FormatAlign::VerticalJustify,
            FormatAlign::VerticalDistributed => xlsx::FormatAlign::VerticalDistributed,
        }
    }
}
