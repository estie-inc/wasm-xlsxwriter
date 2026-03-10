use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `TableFunction` enum defines functions for worksheet table total rows.
///
/// The `TableFunction` enum contains definitions for the standard Excel
/// "SUBTOTAL" functions that are available via the dropdown in the total row of
/// an Excel table. It also supports custom user defined functions or formulas.
#[derive(Debug, Clone, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum TableFunction {
    /// The "total row" option is enable but there is no total function.
    None,
    /// Use the average function as the table total.
    Average,
    /// Use the count function as the table total.
    Count,
    /// Use the count numbers function as the table total.
    CountNumbers,
    /// Use the max function as the table total.
    Max,
    /// Use the min function as the table total.
    Min,
    /// Use the sum function as the table total.
    Sum,
    /// Use the standard deviation function as the table total.
    StdDev,
    /// Use the var function as the table total.
    Var,
    /// Use a custom/user specified function or formula.
    Custom(Formula),
}

impl From<TableFunction> for xlsx::TableFunction {
    fn from(value: TableFunction) -> xlsx::TableFunction {
        match value {
            TableFunction::None => xlsx::TableFunction::None,
            TableFunction::Average => xlsx::TableFunction::Average,
            TableFunction::Count => xlsx::TableFunction::Count,
            TableFunction::CountNumbers => xlsx::TableFunction::CountNumbers,
            TableFunction::Max => xlsx::TableFunction::Max,
            TableFunction::Min => xlsx::TableFunction::Min,
            TableFunction::Sum => xlsx::TableFunction::Sum,
            TableFunction::StdDev => xlsx::TableFunction::StdDev,
            TableFunction::Var => xlsx::TableFunction::Var,
            TableFunction::Custom(v0) => xlsx::TableFunction::Custom(v0),
        }
    }
}
