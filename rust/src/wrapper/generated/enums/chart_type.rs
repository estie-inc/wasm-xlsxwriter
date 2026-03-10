use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ChartType` enum define the type of a {@link Chart} object.
///
/// The main original chart types are supported, see below.
///
/// Support for newer Excel chart types such as Treemap, Sunburst, Box and
/// Whisker, Statistical Histogram, Waterfall, Funnel and Maps is not currently
/// planned since the underlying structure is substantially different from the
/// implemented chart types.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ChartType {
    /// An Area chart type.
    Area,
    /// A stacked Area chart type.
    AreaStacked,
    /// A percent stacked Area chart type.
    AreaPercentStacked,
    /// A Bar (horizontal histogram) chart type.
    Bar,
    /// A stacked Bar chart type.
    BarStacked,
    /// A percent stacked Bar chart type.
    BarPercentStacked,
    /// A Column (vertical histogram) chart type.
    Column,
    /// A stacked Column chart type.
    ColumnStacked,
    /// A percent stacked Column chart type.
    ColumnPercentStacked,
    /// A Doughnut chart type.
    Doughnut,
    /// An Line chart type.
    Line,
    /// A stacked Line chart type.
    LineStacked,
    /// A percent stacked Line chart type.
    LinePercentStacked,
    /// A Pie chart type.
    Pie,
    /// A Radar chart type.
    Radar,
    /// A Radar with markers chart type.
    RadarWithMarkers,
    /// A filled Radar chart type.
    RadarFilled,
    /// A Scatter chart type. Scatter charts, and their variant, are the only
    /// type that have values (as opposed to categories) for their X-Axis. The
    /// default scatter chart in Excel has markers for each point but no
    /// connecting lines.
    Scatter,
    /// A Scatter chart type where the points are connected by straight lines.
    ScatterStraight,
    /// A Scatter chart type where the points have markers and are connected by
    /// straight lines.
    ScatterStraightWithMarkers,
    /// A Scatter chart type where the points are connected by smoothed lines.
    ScatterSmooth,
    /// A Scatter chart type where the points have markers and are connected by
    /// smoothed lines.
    ScatterSmoothWithMarkers,
    /// A Stock chart showing Open-High-Low-Close data. It is also possible to
    /// show High-Low-Close data.
    ///
    /// Note, Volume variants of the Excel stock charts aren't currently
    /// supported but will be in a future release.
    Stock,
}

impl From<ChartType> for xlsx::ChartType {
    fn from(value: ChartType) -> xlsx::ChartType {
        match value {
            ChartType::Area => xlsx::ChartType::Area,
            ChartType::AreaStacked => xlsx::ChartType::AreaStacked,
            ChartType::AreaPercentStacked => xlsx::ChartType::AreaPercentStacked,
            ChartType::Bar => xlsx::ChartType::Bar,
            ChartType::BarStacked => xlsx::ChartType::BarStacked,
            ChartType::BarPercentStacked => xlsx::ChartType::BarPercentStacked,
            ChartType::Column => xlsx::ChartType::Column,
            ChartType::ColumnStacked => xlsx::ChartType::ColumnStacked,
            ChartType::ColumnPercentStacked => xlsx::ChartType::ColumnPercentStacked,
            ChartType::Doughnut => xlsx::ChartType::Doughnut,
            ChartType::Line => xlsx::ChartType::Line,
            ChartType::LineStacked => xlsx::ChartType::LineStacked,
            ChartType::LinePercentStacked => xlsx::ChartType::LinePercentStacked,
            ChartType::Pie => xlsx::ChartType::Pie,
            ChartType::Radar => xlsx::ChartType::Radar,
            ChartType::RadarWithMarkers => xlsx::ChartType::RadarWithMarkers,
            ChartType::RadarFilled => xlsx::ChartType::RadarFilled,
            ChartType::Scatter => xlsx::ChartType::Scatter,
            ChartType::ScatterStraight => xlsx::ChartType::ScatterStraight,
            ChartType::ScatterStraightWithMarkers => xlsx::ChartType::ScatterStraightWithMarkers,
            ChartType::ScatterSmooth => xlsx::ChartType::ScatterSmooth,
            ChartType::ScatterSmoothWithMarkers => xlsx::ChartType::ScatterSmoothWithMarkers,
            ChartType::Stock => xlsx::ChartType::Stock,
        }
    }
}
