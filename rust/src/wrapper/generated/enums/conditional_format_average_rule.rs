use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ConditionalFormatAverageRule` enum defines the conditional format
/// criteria for {@link ConditionalFormatCell}.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ConditionalFormatAverageRule {
    /// Show the conditional format for cells above the average for the range.
    /// This is the default.
    AboveAverage,
    /// Show the conditional format for cells below the average for the range.
    BelowAverage,
    /// Show the conditional format for cells above or equal to the average for
    /// the range.
    EqualOrAboveAverage,
    /// Show the conditional format for cells below or equal to the average for
    /// the range.
    EqualOrBelowAverage,
    /// Show the conditional format for cells 1 standard deviation above the
    /// average for the range.
    OneStandardDeviationAbove,
    /// Show the conditional format for cells 1 standard deviation below the
    /// average for the range.
    OneStandardDeviationBelow,
    /// Show the conditional format for cells 2 standard deviation above the
    /// average for the range.
    TwoStandardDeviationsAbove,
    /// Show the conditional format for cells 2 standard deviation below the
    /// average for the range.
    TwoStandardDeviationsBelow,
    /// Show the conditional format for cells 3 standard deviation above the
    /// average for the range.
    ThreeStandardDeviationsAbove,
    /// Show the conditional format for cells 3 standard deviation below the
    /// average for the range.
    ThreeStandardDeviationsBelow,
}

impl From<ConditionalFormatAverageRule> for xlsx::ConditionalFormatAverageRule {
    fn from(value: ConditionalFormatAverageRule) -> xlsx::ConditionalFormatAverageRule {
        match value {
            ConditionalFormatAverageRule::AboveAverage => {
                xlsx::ConditionalFormatAverageRule::AboveAverage
            }
            ConditionalFormatAverageRule::BelowAverage => {
                xlsx::ConditionalFormatAverageRule::BelowAverage
            }
            ConditionalFormatAverageRule::EqualOrAboveAverage => {
                xlsx::ConditionalFormatAverageRule::EqualOrAboveAverage
            }
            ConditionalFormatAverageRule::EqualOrBelowAverage => {
                xlsx::ConditionalFormatAverageRule::EqualOrBelowAverage
            }
            ConditionalFormatAverageRule::OneStandardDeviationAbove => {
                xlsx::ConditionalFormatAverageRule::OneStandardDeviationAbove
            }
            ConditionalFormatAverageRule::OneStandardDeviationBelow => {
                xlsx::ConditionalFormatAverageRule::OneStandardDeviationBelow
            }
            ConditionalFormatAverageRule::TwoStandardDeviationsAbove => {
                xlsx::ConditionalFormatAverageRule::TwoStandardDeviationsAbove
            }
            ConditionalFormatAverageRule::TwoStandardDeviationsBelow => {
                xlsx::ConditionalFormatAverageRule::TwoStandardDeviationsBelow
            }
            ConditionalFormatAverageRule::ThreeStandardDeviationsAbove => {
                xlsx::ConditionalFormatAverageRule::ThreeStandardDeviationsAbove
            }
            ConditionalFormatAverageRule::ThreeStandardDeviationsBelow => {
                xlsx::ConditionalFormatAverageRule::ThreeStandardDeviationsBelow
            }
        }
    }
}
