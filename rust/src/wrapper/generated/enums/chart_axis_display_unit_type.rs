use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ChartAxisDisplayUnitType` enum defines the Chart axis date display
/// unit types.
///
/// Define the display unit type for chart axes such as "Thousands" or
/// "Millions".
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ChartAxisDisplayUnitType {
    /// Don't display any units for the axis values, the default.
    None,
    /// Display the axis values in units of Hundreds.
    Hundreds,
    /// Display the axis values in units of Thousands.
    Thousands,
    /// Display the axis values in units of Ten Thousands.
    TenThousands,
    /// Display the axis values in units of Hundred Thousands.
    HundredThousands,
    /// Display the axis values in units of Millions.
    Millions,
    /// Display the axis values in units of Ten Millions.
    TenMillions,
    /// Display the axis values in units of Hundred Millions.
    HundredMillions,
    /// Display the axis values in units of Billions.
    Billions,
    /// Display the axis values in units of Trillions.
    Trillions,
}

impl From<ChartAxisDisplayUnitType> for xlsx::ChartAxisDisplayUnitType {
    fn from(value: ChartAxisDisplayUnitType) -> xlsx::ChartAxisDisplayUnitType {
        match value {
            ChartAxisDisplayUnitType::None => xlsx::ChartAxisDisplayUnitType::None,
            ChartAxisDisplayUnitType::Hundreds => xlsx::ChartAxisDisplayUnitType::Hundreds,
            ChartAxisDisplayUnitType::Thousands => xlsx::ChartAxisDisplayUnitType::Thousands,
            ChartAxisDisplayUnitType::TenThousands => xlsx::ChartAxisDisplayUnitType::TenThousands,
            ChartAxisDisplayUnitType::HundredThousands => {
                xlsx::ChartAxisDisplayUnitType::HundredThousands
            }
            ChartAxisDisplayUnitType::Millions => xlsx::ChartAxisDisplayUnitType::Millions,
            ChartAxisDisplayUnitType::TenMillions => xlsx::ChartAxisDisplayUnitType::TenMillions,
            ChartAxisDisplayUnitType::HundredMillions => {
                xlsx::ChartAxisDisplayUnitType::HundredMillions
            }
            ChartAxisDisplayUnitType::Billions => xlsx::ChartAxisDisplayUnitType::Billions,
            ChartAxisDisplayUnitType::Trillions => xlsx::ChartAxisDisplayUnitType::Trillions,
        }
    }
}
