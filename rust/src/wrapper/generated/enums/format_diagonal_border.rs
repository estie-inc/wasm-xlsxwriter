use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `FormatDiagonalBorder` enum defines Format diagonal border types.
///
/// This is used with the Format.setBorderDiagonal() method.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify, Default)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum FormatDiagonalBorder {
    /// The default/automatic format for an Excel font.
    #[default]
    None,
    /// Cell diagonal border from bottom left to top right.
    BorderUp,
    /// Cell diagonal border from top left to bottom right.
    BorderDown,
    /// Cell diagonal border in both directions.
    BorderUpDown,
}

impl From<FormatDiagonalBorder> for xlsx::FormatDiagonalBorder {
    fn from(value: FormatDiagonalBorder) -> xlsx::FormatDiagonalBorder {
        match value {
            FormatDiagonalBorder::None => xlsx::FormatDiagonalBorder::None,
            FormatDiagonalBorder::BorderUp => xlsx::FormatDiagonalBorder::BorderUp,
            FormatDiagonalBorder::BorderDown => xlsx::FormatDiagonalBorder::BorderDown,
            FormatDiagonalBorder::BorderUpDown => xlsx::FormatDiagonalBorder::BorderUpDown,
        }
    }
}
