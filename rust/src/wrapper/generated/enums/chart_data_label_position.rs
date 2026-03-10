use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ChartDataLabelPosition` enum defines the Chart data label positions.
///
/// In Excel the available data label positions vary for different chart
/// types. The available, and default, positions are:
///
/// | Position      | Line, Scatter | Bar, Column   | Pie, Doughnut | Area, Radar   |
/// | :------------ | :------------ | :------------ | :------------ | :------------ |
/// | `Center`      | Yes           | Yes           | Yes           | Yes (default) |
/// | `Right`       | Yes (default) |               |               |               |
/// | `Left`        | Yes           |               |               |               |
/// | `Above`       | Yes           |               |               |               |
/// | `Below`       | Yes           |               |               |               |
/// | `InsideBase`  |               | Yes           |               |               |
/// | `InsideEnd`   |               | Yes           | Yes           |               |
/// | `OutsideEnd`  |               | Yes (default) | Yes           |               |
/// | `BestFit`     |               |               | Yes (default) |               |
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ChartDataLabelPosition {
    /// Series data label position: Default position.
    Default,
    /// Series data label position: Center.
    Center,
    /// Series data label position: Right.
    Right,
    /// Series data label position: Left.
    Left,
    /// Series data label position: Above.
    Above,
    /// Series data label position: Below.
    Below,
    /// Series data label position: Inside base.
    InsideBase,
    /// Series data label position: Inside end.
    InsideEnd,
    /// Series data label position: Outside end.
    OutsideEnd,
    /// Series data label position: Best fit.
    BestFit,
}

impl From<ChartDataLabelPosition> for xlsx::ChartDataLabelPosition {
    fn from(value: ChartDataLabelPosition) -> xlsx::ChartDataLabelPosition {
        match value {
            ChartDataLabelPosition::Default => xlsx::ChartDataLabelPosition::Default,
            ChartDataLabelPosition::Center => xlsx::ChartDataLabelPosition::Center,
            ChartDataLabelPosition::Right => xlsx::ChartDataLabelPosition::Right,
            ChartDataLabelPosition::Left => xlsx::ChartDataLabelPosition::Left,
            ChartDataLabelPosition::Above => xlsx::ChartDataLabelPosition::Above,
            ChartDataLabelPosition::Below => xlsx::ChartDataLabelPosition::Below,
            ChartDataLabelPosition::InsideBase => xlsx::ChartDataLabelPosition::InsideBase,
            ChartDataLabelPosition::InsideEnd => xlsx::ChartDataLabelPosition::InsideEnd,
            ChartDataLabelPosition::OutsideEnd => xlsx::ChartDataLabelPosition::OutsideEnd,
            ChartDataLabelPosition::BestFit => xlsx::ChartDataLabelPosition::BestFit,
        }
    }
}
