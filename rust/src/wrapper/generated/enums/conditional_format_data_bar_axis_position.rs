use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ConditionalFormatDataBarAxisPosition` enum defines the conditional
/// format axis positions for ConditionalFormatDataBar.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ConditionalFormatDataBarAxisPosition {
    /// The axis is set automatically depending on whether the data contains
    /// negative values. This is the default.
    Automatic,
    /// The axis is set at the midpoint. This is the automatic option for ranges
    /// with negative values.
    Midpoint,
    /// Turn the axis off.
    None,
}

impl From<ConditionalFormatDataBarAxisPosition> for xlsx::ConditionalFormatDataBarAxisPosition {
    fn from(
        value: ConditionalFormatDataBarAxisPosition,
    ) -> xlsx::ConditionalFormatDataBarAxisPosition {
        match value {
            ConditionalFormatDataBarAxisPosition::Automatic => {
                xlsx::ConditionalFormatDataBarAxisPosition::Automatic
            }
            ConditionalFormatDataBarAxisPosition::Midpoint => {
                xlsx::ConditionalFormatDataBarAxisPosition::Midpoint
            }
            ConditionalFormatDataBarAxisPosition::None => {
                xlsx::ConditionalFormatDataBarAxisPosition::None
            }
        }
    }
}
