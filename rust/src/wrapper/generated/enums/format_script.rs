use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `FormatScript` enum defines the {@link Format} font superscript and subscript
/// properties.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify, Default)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum FormatScript {
    /// The default/automatic format for an Excel font.
    #[default]
    None,
    /// The cell text is superscripted.
    Superscript,
    /// The cell text is subscripted.
    Subscript,
}

impl From<FormatScript> for xlsx::FormatScript {
    fn from(value: FormatScript) -> xlsx::FormatScript {
        match value {
            FormatScript::None => xlsx::FormatScript::None,
            FormatScript::Superscript => xlsx::FormatScript::Superscript,
            FormatScript::Subscript => xlsx::FormatScript::Subscript,
        }
    }
}
