use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

/// The `ChartType` enum define the type of a {@link Chart} object.
///
/// The main original chart types are supported, see below.
///
/// Support for newer Excel chart types such as Treemap, Sunburst, Box and
/// Whisker, Statistical Histogram, Waterfall, Funnel and Maps is not currently
/// planned since the underlying structure is substantially different from the
/// implemented chart types.
///
#[derive(Clone, Copy, PartialEq, Eq)]
#[wasm_bindgen]
pub enum ChartType {
    /// An Area chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_area.png">
    Area,
    /// A stacked Area chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_area_stacked.png">
    AreaStacked,
    /// A percent stacked Area chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_area_percent_stacked.png">
    AreaPercentStacked,
    /// A Bar (horizontal histogram) chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_bar.png">
    Bar,
    /// A stacked Bar chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_bar_stacked.png">
    BarStacked,
    /// A percent stacked Bar chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_bar_percent_stacked.png">
    BarPercentStacked,
    /// A Column (vertical histogram) chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_column.png">
    Column,
    /// A stacked Column chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_column_stacked.png">
    ColumnStacked,
    /// A percent stacked Column chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_column_percent_stacked.png">
    ColumnPercentStacked,
    /// A Doughnut chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_doughnut.png">
    Doughnut,
    /// An Line chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_line.png">
    Line,
    /// A stacked Line chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_line_stacked.png">
    LineStacked,
    /// A percent stacked Line chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_line_percent_stacked.png">
    LinePercentStacked,
    /// A Pie chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_pie.png">
    Pie,
    /// A Radar chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_radar.png">
    Radar,
    /// A Radar with markers chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_radar_with_markers.png">
    RadarWithMarkers,
    /// A filled Radar chart type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_radar_filled.png">
    RadarFilled,
    /// A Scatter chart type. Scatter charts, and their variant, are the only
    /// type that have values (as opposed to categories) for their X-Axis. The
    /// default scatter chart in Excel has markers for each point but no
    /// connecting lines.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_scatter.png">
    Scatter,
    /// A Scatter chart type where the points are connected by straight lines.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_scatter_straight.png">
    ScatterStraight,
    /// A Scatter chart type where the points have markers and are connected by
    /// straight lines.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_scatter_straight_with_markers.png">
    ScatterStraightWithMarkers,
    /// A Scatter chart type where the points are connected by smoothed lines.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_scatter_smooth.png">
    ScatterSmooth,
    /// A Scatter chart type where the points have markers and are connected by
    /// smoothed lines.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_scatter_smooth_with_markers.png">
    ScatterSmoothWithMarkers,
    /// A Stock chart showing Open-High-Low-Close data. It is also possible to
    /// show High-Low-Close data.
    ///
    /// Note, Volume variants of the Excel stock charts aren't currently
    /// supported but will be in a future release.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_type_stock.png">
    Stock,
}

impl From<ChartType> for xlsx::ChartType {
    fn from(chart_type: ChartType) -> xlsx::ChartType {
        match chart_type {
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
