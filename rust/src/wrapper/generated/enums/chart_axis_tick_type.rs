use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ChartAxisTickType` enum defines the Chart axis tick types.
///
/// Excel supports 4 types of tick position:
///
/// - None
/// - Inside only
/// - Outside only
/// - Cross - inside and outside
///
/// Used in conjunction with ChartAxis.setMajorTickType() and
/// ChartAxis.setMinorTickType().
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ChartAxisTickType {
    /// No tick mark for the axis.
    None,
    /// The tick mark is inside the axis only.
    Inside,
    /// The tick mark is outside the axis only.
    Outside,
    /// The tick mark crosses inside and outside the axis.
    Cross,
}

impl From<ChartAxisTickType> for xlsx::ChartAxisTickType {
    fn from(value: ChartAxisTickType) -> xlsx::ChartAxisTickType {
        match value {
            ChartAxisTickType::None => xlsx::ChartAxisTickType::None,
            ChartAxisTickType::Inside => xlsx::ChartAxisTickType::Inside,
            ChartAxisTickType::Outside => xlsx::ChartAxisTickType::Outside,
            ChartAxisTickType::Cross => xlsx::ChartAxisTickType::Cross,
        }
    }
}
