use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `DataValidationRule` enum defines the data validation rule for
/// DataValidation.
#[derive(Debug, Clone, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum DataValidationRule {
    /// Restrict cell input to values that are equal to the target value.
    EqualTo(T),
    /// Restrict cell input to values that are not equal to the target value.
    NotEqualTo(T),
    /// Restrict cell input to values that are greater than the target value.
    GreaterThan(T),
    /// Restrict cell input to values that are greater than or equal to the
    /// target value.
    GreaterThanOrEqualTo(T),
    /// Restrict cell input to values that are less than the target value.
    LessThan(T),
    /// Restrict cell input to values that are less than or equal to the target
    /// value.
    LessThanOrEqualTo(T),
    /// Restrict cell input to values that are between the target values.
    Between(T, T),
    /// Restrict cell input to values that are not between the target values.
    NotBetween(T, T),
}

impl From<DataValidationRule> for xlsx::DataValidationRule {
    fn from(value: DataValidationRule) -> xlsx::DataValidationRule {
        match value {
            DataValidationRule::EqualTo(v0) => xlsx::DataValidationRule::EqualTo(v0),
            DataValidationRule::NotEqualTo(v0) => xlsx::DataValidationRule::NotEqualTo(v0),
            DataValidationRule::GreaterThan(v0) => xlsx::DataValidationRule::GreaterThan(v0),
            DataValidationRule::GreaterThanOrEqualTo(v0) => {
                xlsx::DataValidationRule::GreaterThanOrEqualTo(v0)
            }
            DataValidationRule::LessThan(v0) => xlsx::DataValidationRule::LessThan(v0),
            DataValidationRule::LessThanOrEqualTo(v0) => {
                xlsx::DataValidationRule::LessThanOrEqualTo(v0)
            }
            DataValidationRule::Between(v0, v1) => xlsx::DataValidationRule::Between(v0, v1),
            DataValidationRule::NotBetween(v0, v1) => xlsx::DataValidationRule::NotBetween(v0, v1),
        }
    }
}
