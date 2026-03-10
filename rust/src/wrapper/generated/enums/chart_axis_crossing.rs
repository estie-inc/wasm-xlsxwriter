use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ChartAxisCrossing` enum defines the ChartAxis crossing point for
/// the opposite axis.
///
/// By default Excel sets chart axes to cross at 0. If required you can use
/// ChartAxis.setCrossing() and ChartAxisCrossing to define another
/// point where the opposite axis will cross the current axis.
#[derive(Debug, Clone, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ChartAxisCrossing {
    /// The axis crossing is at the default value which is generally zero. This
    /// is the default.
    Automatic,
    /// The axis crossing is at the minimum value for the axis.
    Min,
    /// The axis crossing is at the maximum value for the axis.
    Max,
    /// The axis crossing is at a category index number.
    ///
    /// This is for Category style axes only. For example say you are plotting 4
    /// categories on the X-axis ("North", "South", "East", "West"). By setting
    /// the category number as `CategoryNumber(3)` the Y-axis will cross at
    /// "East".
    ///
    /// See [Chart Value and Category
    /// Axes](crate::chart#chart-value-and-category-axes) for an
    /// explanation of the difference between Value and Category axes in Excel.
    CategoryNumber(u32),
    /// The axis crossing is at a value.
    ///
    /// This is for Value and Date style axes only.
    ///
    /// See [Chart Value and Category
    /// Axes](crate::chart#chart-value-and-category-axes) for an
    /// explanation of the difference between Value and Category axes in Excel.
    AxisValue(f64),
}

impl From<ChartAxisCrossing> for xlsx::ChartAxisCrossing {
    fn from(value: ChartAxisCrossing) -> xlsx::ChartAxisCrossing {
        match value {
            ChartAxisCrossing::Automatic => xlsx::ChartAxisCrossing::Automatic,
            ChartAxisCrossing::Min => xlsx::ChartAxisCrossing::Min,
            ChartAxisCrossing::Max => xlsx::ChartAxisCrossing::Max,
            ChartAxisCrossing::CategoryNumber(v0) => xlsx::ChartAxisCrossing::CategoryNumber(v0),
            ChartAxisCrossing::AxisValue(v0) => xlsx::ChartAxisCrossing::AxisValue(v0),
        }
    }
}
