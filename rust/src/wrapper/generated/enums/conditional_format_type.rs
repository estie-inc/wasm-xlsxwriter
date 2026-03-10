use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ConditionalFormatType` enum defines the conditional format type
/// for ConditionalFormat2ColorScale, ConditionalFormat3ColorScale and
/// ConditionalFormatDataBar.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ConditionalFormatType {
    /// Set the color scale/data bar to use the minimum or maximum value in the range. This is the
    /// default for data bars.
    Automatic,
    /// Set the color scale/data bar to use the minimum value in the range. This is the
    /// default for the minimum value in color scales.
    Lowest,
    /// Set the color scale/data bar to use a number value other than the
    /// maximum/minimum.
    Number,
    /// Set the color scale/data bar to use a percentage. This must be in the range
    /// 0-100.
    Percent,
    /// Set the color scale/data bar to use a formula value.
    Formula,
    /// Set the color scale/data bar to use a percentile. This must be in the range
    /// 0-100.
    Percentile,
    /// Set the color scale/data bar to use the maximum value in the range. This is the
    /// default for the maximum value in value in color scales.
    Highest,
}

impl From<ConditionalFormatType> for xlsx::ConditionalFormatType {
    fn from(value: ConditionalFormatType) -> xlsx::ConditionalFormatType {
        match value {
            ConditionalFormatType::Automatic => xlsx::ConditionalFormatType::Automatic,
            ConditionalFormatType::Lowest => xlsx::ConditionalFormatType::Lowest,
            ConditionalFormatType::Number => xlsx::ConditionalFormatType::Number,
            ConditionalFormatType::Percent => xlsx::ConditionalFormatType::Percent,
            ConditionalFormatType::Formula => xlsx::ConditionalFormatType::Formula,
            ConditionalFormatType::Percentile => xlsx::ConditionalFormatType::Percentile,
            ConditionalFormatType::Highest => xlsx::ConditionalFormatType::Highest,
        }
    }
}
