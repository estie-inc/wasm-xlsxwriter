use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `IgnoreError` enum defines the Excel cell error types that can be
/// ignored.
///
/// The equivalent options in Excel are:
///
/// <img src="https://rustxlsxwriter.github.io/images/ignore_errors_dialog.png">
///
/// Note, some of the items shown in the above dialog such as "Misleading Number
/// Formats" aren't saved in the output file by Excel and can't be turned off
/// permanently.
///
/// The enum values are used with the Worksheet.ignoreError() and
/// Worksheet.ignoreErrorRange() methods.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum IgnoreError {
    /// Ignore errors/warnings for numbers stored as text.
    NumberStoredAsText,
    /// Ignore errors/warnings for formula evaluation errors (such as divide by
    /// zero).
    FormulaError,
    /// Ignore errors/warnings for formulas that differ from surrounding
    /// formulas.
    FormulaDiffers,
    /// Ignore errors/warnings for formulas that refer to empty cells.
    FormulaRefersToEmptyCells,
    /// Ignore errors/warnings for formulas that omit cells in a range.
    FormulaOmitsCells,
    /// Ignore errors/warnings for cells in a table that do not comply with
    /// applicable data validation rules.
    DataValidationError,
    /// Ignore errors/warnings for formulas that contain a two digit text
    /// representation of a year.
    TwoDigitTextYear,
    /// Ignore  errors/warnings for unlocked cells that contain formulas.
    UnlockedCellsWithFormula,
    /// Ignore errors/warnings for cell formulas that differ from the column
    /// formula.
    InconsistentColumnFormula,
}

impl From<IgnoreError> for xlsx::IgnoreError {
    fn from(value: IgnoreError) -> xlsx::IgnoreError {
        match value {
            IgnoreError::NumberStoredAsText => xlsx::IgnoreError::NumberStoredAsText,
            IgnoreError::FormulaError => xlsx::IgnoreError::FormulaError,
            IgnoreError::FormulaDiffers => xlsx::IgnoreError::FormulaDiffers,
            IgnoreError::FormulaRefersToEmptyCells => xlsx::IgnoreError::FormulaRefersToEmptyCells,
            IgnoreError::FormulaOmitsCells => xlsx::IgnoreError::FormulaOmitsCells,
            IgnoreError::DataValidationError => xlsx::IgnoreError::DataValidationError,
            IgnoreError::TwoDigitTextYear => xlsx::IgnoreError::TwoDigitTextYear,
            IgnoreError::UnlockedCellsWithFormula => xlsx::IgnoreError::UnlockedCellsWithFormula,
            IgnoreError::InconsistentColumnFormula => xlsx::IgnoreError::InconsistentColumnFormula,
        }
    }
}
