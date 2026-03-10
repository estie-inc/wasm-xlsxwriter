use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ConditionalFormatTextRule` enum defines the conditional format
/// criteria for {@link ConditionalFormatText}.
#[derive(Debug, Clone, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ConditionalFormatTextRule {
    /// Show the conditional format for text that contains to the target string.
    Contains(String),
    /// Show the conditional format for text that do not contain to the target string.
    DoesNotContain(String),
    /// Show the conditional format for text that begins with the target string.
    BeginsWith(String),
    /// Show the conditional format for text that ends with the target string.
    EndsWith(String),
}

impl From<ConditionalFormatTextRule> for xlsx::ConditionalFormatTextRule {
    fn from(value: ConditionalFormatTextRule) -> xlsx::ConditionalFormatTextRule {
        match value {
            ConditionalFormatTextRule::Contains(v0) => {
                xlsx::ConditionalFormatTextRule::Contains(v0)
            }
            ConditionalFormatTextRule::DoesNotContain(v0) => {
                xlsx::ConditionalFormatTextRule::DoesNotContain(v0)
            }
            ConditionalFormatTextRule::BeginsWith(v0) => {
                xlsx::ConditionalFormatTextRule::BeginsWith(v0)
            }
            ConditionalFormatTextRule::EndsWith(v0) => {
                xlsx::ConditionalFormatTextRule::EndsWith(v0)
            }
        }
    }
}
