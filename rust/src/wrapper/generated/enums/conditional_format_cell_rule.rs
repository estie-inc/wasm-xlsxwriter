use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ConditionalFormatCellRule` enum defines the conditional format rule for
/// ConditionalFormatCell.
#[derive(Debug, Clone, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ConditionalFormatCellRule {
    /// Show the conditional format for cells that are equal to the target value.
    EqualTo(T),
    /// Show the conditional format for cells that are not equal to the target value.
    NotEqualTo(T),
    /// Show the conditional format for cells that are greater than the target value.
    GreaterThan(T),
    /// Show the conditional format for cells that are greater than or equal to the target value.
    GreaterThanOrEqualTo(T),
    /// Show the conditional format for cells that are less than the target value.
    LessThan(T),
    /// Show the conditional format for cells that are less than or equal to the target value.
    LessThanOrEqualTo(T),
    /// Show the conditional format for cells that are between the target values.
    Between(T, T),
    /// Show the conditional format for cells that are not between the target values.
    NotBetween(T, T),
}

impl From<ConditionalFormatCellRule> for xlsx::ConditionalFormatCellRule {
    fn from(value: ConditionalFormatCellRule) -> xlsx::ConditionalFormatCellRule {
        match value {
            ConditionalFormatCellRule::EqualTo(v0) => xlsx::ConditionalFormatCellRule::EqualTo(v0),
            ConditionalFormatCellRule::NotEqualTo(v0) => {
                xlsx::ConditionalFormatCellRule::NotEqualTo(v0)
            }
            ConditionalFormatCellRule::GreaterThan(v0) => {
                xlsx::ConditionalFormatCellRule::GreaterThan(v0)
            }
            ConditionalFormatCellRule::GreaterThanOrEqualTo(v0) => {
                xlsx::ConditionalFormatCellRule::GreaterThanOrEqualTo(v0)
            }
            ConditionalFormatCellRule::LessThan(v0) => {
                xlsx::ConditionalFormatCellRule::LessThan(v0)
            }
            ConditionalFormatCellRule::LessThanOrEqualTo(v0) => {
                xlsx::ConditionalFormatCellRule::LessThanOrEqualTo(v0)
            }
            ConditionalFormatCellRule::Between(v0, v1) => {
                xlsx::ConditionalFormatCellRule::Between(v0, v1)
            }
            ConditionalFormatCellRule::NotBetween(v0, v1) => {
                xlsx::ConditionalFormatCellRule::NotBetween(v0, v1)
            }
        }
    }
}
