use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ConditionalFormatTopRule` enum defines the conditional format rule for
/// ConditionalFormatCell.
#[derive(Debug, Clone, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ConditionalFormatTopRule {
    /// Show the conditional format for cells that are in the top X.
    Top(u16),
    /// Show the conditional format for cells that are in the bottom X.
    Bottom(u16),
    /// Show the conditional format for cells that are in the top X%.
    TopPercent(u16),
    /// Show the conditional format for cells that are in the bottom X%.
    BottomPercent(u16),
}

impl From<ConditionalFormatTopRule> for xlsx::ConditionalFormatTopRule {
    fn from(value: ConditionalFormatTopRule) -> xlsx::ConditionalFormatTopRule {
        match value {
            ConditionalFormatTopRule::Top(v0) => xlsx::ConditionalFormatTopRule::Top(v0),
            ConditionalFormatTopRule::Bottom(v0) => xlsx::ConditionalFormatTopRule::Bottom(v0),
            ConditionalFormatTopRule::TopPercent(v0) => {
                xlsx::ConditionalFormatTopRule::TopPercent(v0)
            }
            ConditionalFormatTopRule::BottomPercent(v0) => {
                xlsx::ConditionalFormatTopRule::BottomPercent(v0)
            }
        }
    }
}
