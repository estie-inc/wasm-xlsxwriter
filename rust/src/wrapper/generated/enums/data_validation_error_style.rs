use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `DataValidationErrorStyle` enum defines the type of error dialog that is
/// shown when there is and error in a data validation.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum DataValidationErrorStyle {
    /// Show a "Stop" dialog. This is the default.
    Stop,
    /// Show a "Warning" dialog.
    Warning,
    /// Show an "Information" dialog.
    Information,
}

impl From<DataValidationErrorStyle> for xlsx::DataValidationErrorStyle {
    fn from(value: DataValidationErrorStyle) -> xlsx::DataValidationErrorStyle {
        match value {
            DataValidationErrorStyle::Stop => xlsx::DataValidationErrorStyle::Stop,
            DataValidationErrorStyle::Warning => xlsx::DataValidationErrorStyle::Warning,
            DataValidationErrorStyle::Information => xlsx::DataValidationErrorStyle::Information,
        }
    }
}
