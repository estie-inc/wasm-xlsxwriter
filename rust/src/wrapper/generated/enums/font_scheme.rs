use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `FontScheme` enum defines the font scheme properties of a {@link Format}.
/// These relate to whether the font is part of the theme or is a custom font.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify, Default)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum FontScheme {
    /// This is used to indicate a theme font that is used for body text.
    Body,
    /// This is used to indicate a theme font that is used for headings. Note,
    /// this is rarely, if ever, required when using Excel.
    Headings,
    /// This is used to indicate a custom font that is not part of the theme.
    /// For example, this can also be used to create a Calibri 11 font that is
    /// not part of the default theme.
    #[default]
    None,
}

impl From<FontScheme> for xlsx::FontScheme {
    fn from(value: FontScheme) -> xlsx::FontScheme {
        match value {
            FontScheme::Body => xlsx::FontScheme::Body,
            FontScheme::Headings => xlsx::FontScheme::Headings,
            FontScheme::None => xlsx::FontScheme::None,
        }
    }
}
