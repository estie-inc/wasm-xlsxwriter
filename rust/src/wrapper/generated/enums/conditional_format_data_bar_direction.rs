use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ConditionalFormatDataBarDirection` enum defines the conditional format
/// directions for ConditionalFormatDataBar. This is used to set the data
/// bar conditional format direction to "Right to left", "Left to right" or
/// "Context" (the default) in conjunction with
/// ConditionalFormatDataBar.setDirection().
///
/// # Parameters
///
/// - `direction`: A ConditionalFormatDataBarDirection enum value.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ConditionalFormatDataBarDirection {
    /// The bars go "Right to left" or "Left to right" depending on the context.
    /// This is the default.
    Context,
    /// The bars go "Left to right".
    LeftToRight,
    /// The bars go "Right to left".
    RightToLeft,
}

impl From<ConditionalFormatDataBarDirection> for xlsx::ConditionalFormatDataBarDirection {
    fn from(value: ConditionalFormatDataBarDirection) -> xlsx::ConditionalFormatDataBarDirection {
        match value {
            ConditionalFormatDataBarDirection::Context => {
                xlsx::ConditionalFormatDataBarDirection::Context
            }
            ConditionalFormatDataBarDirection::LeftToRight => {
                xlsx::ConditionalFormatDataBarDirection::LeftToRight
            }
            ConditionalFormatDataBarDirection::RightToLeft => {
                xlsx::ConditionalFormatDataBarDirection::RightToLeft
            }
        }
    }
}
