use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `FormatUnderline` enum defines the font underline type in a {@link Format}.
///
/// The difference between a normal underline and an "accounting" underline is
/// that a normal underline only underlines the text/number in a cell whereas an
/// accounting underline underlines the entire cell width.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify, Default)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum FormatUnderline {
    /// The default/automatic underline for an Excel font.
    #[default]
    None,
    /// A single underline under the text/number in a cell.
    Single,
    /// A double underline under the text/number in a cell.
    Double,
    /// A single accounting style underline under the entire cell.
    SingleAccounting,
    /// A double accounting style underline under the entire cell.
    DoubleAccounting,
}

impl From<FormatUnderline> for xlsx::FormatUnderline {
    fn from(value: FormatUnderline) -> xlsx::FormatUnderline {
        match value {
            FormatUnderline::None => xlsx::FormatUnderline::None,
            FormatUnderline::Single => xlsx::FormatUnderline::Single,
            FormatUnderline::Double => xlsx::FormatUnderline::Double,
            FormatUnderline::SingleAccounting => xlsx::FormatUnderline::SingleAccounting,
            FormatUnderline::DoubleAccounting => xlsx::FormatUnderline::DoubleAccounting,
        }
    }
}
