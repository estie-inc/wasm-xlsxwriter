use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ChartAxisDateUnitType` enum defines the {@link Chart} axis date unit types.
///
/// Define the unit type for the major or minor unit in a Chart Date axis.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ChartAxisDateUnitType {
    /// The major or minor unit is expressed in days.
    Days,
    /// The major or minor unit is expressed in months.
    Months,
    /// The major or minor unit is expressed in years.
    Years,
}

impl From<ChartAxisDateUnitType> for xlsx::ChartAxisDateUnitType {
    fn from(value: ChartAxisDateUnitType) -> xlsx::ChartAxisDateUnitType {
        match value {
            ChartAxisDateUnitType::Days => xlsx::ChartAxisDateUnitType::Days,
            ChartAxisDateUnitType::Months => xlsx::ChartAxisDateUnitType::Months,
            ChartAxisDateUnitType::Years => xlsx::ChartAxisDateUnitType::Years,
        }
    }
}
