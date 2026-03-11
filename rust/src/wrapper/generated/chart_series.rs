use crate::wrapper::ChartDataLabel;
use crate::wrapper::ChartErrorBars;
use crate::wrapper::ChartMarker;
use crate::wrapper::ChartPoint;
use crate::wrapper::ChartTrendline;
use crate::wrapper::Color;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ChartSeries` struct represents a chart series.
///
/// A chart in Excel can contain one of more data series. The `ChartSeries`
/// struct represents the Category and Value ranges, and the formatting and
/// options for the chart series.
///
/// It is used in conjunction with the {@link Chart} struct.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartSeries {
    pub(crate) parent: Arc<Mutex<xlsx::Chart>>,
}

#[wasm_bindgen]
impl ChartSeries {
    /// Plot the chart series on the secondary axis.
    ///
    /// It is possible to add a secondary axis of the same type to a chart by
    /// setting the `secondary_axis` property of the series. See Chart Secondary
    /// Axes
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setSecondaryAxis", skip_jsdoc)]
    pub fn set_secondary_axis(&self, enable: bool) -> ChartSeries {
        let mut lock = self.parent.lock().unwrap();
        lock.add_series().set_secondary_axis(enable);
        ChartSeries {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the markers for a chart series.
    ///
    /// Set the markers and marker properties for a data series using a
    /// {@link ChartMarker} instance. In general only Line, Scatter and Radar chart
    /// support markers.
    ///
    /// # Parameters
    ///
    /// `marker`: A {@link ChartMarker} instance.
    #[wasm_bindgen(js_name = "setMarker", skip_jsdoc)]
    pub fn set_marker(&self, marker: ChartMarker) -> ChartSeries {
        let mut lock = self.parent.lock().unwrap();
        lock.add_series().set_marker(&marker.inner);
        ChartSeries {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the data labels for a chart series.
    ///
    /// Set the data labels and marker properties for a data series using a
    /// {@link ChartDataLabel} instance.
    ///
    /// # Parameters
    ///
    /// `data_label`: A {@link ChartDataLabel} instance.
    #[wasm_bindgen(js_name = "setDataLabel", skip_jsdoc)]
    pub fn set_data_label(&self, data_label: ChartDataLabel) -> ChartSeries {
        let mut lock = self.parent.lock().unwrap();
        lock.add_series().set_data_label(&data_label.inner);
        ChartSeries {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set custom data labels for a data series.
    ///
    /// The {@link ChartSeries#setDataLabel} method sets the data label
    /// properties for every label in a series but it is occasionally required
    /// to set separate properties for individual data labels, or set a custom
    /// display value, or format or hide some of the labels. This can be
    /// achieved with the `set_custom_data_labels()` method, see the examples
    /// below.
    ///
    /// # Parameters
    ///
    /// `data_labels`: A slice of {@link ChartDataLabel} objects.
    #[wasm_bindgen(js_name = "setCustomDataLabels", skip_jsdoc)]
    pub fn set_custom_data_labels(&self, data_labels: Vec<ChartDataLabel>) -> ChartSeries {
        let mut lock = self.parent.lock().unwrap();
        lock.add_series().set_custom_data_labels(
            &data_labels
                .iter()
                .map(|x| x.inner.clone())
                .collect::<Vec<_>>(),
        );
        ChartSeries {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the formatting and properties for points in a chart series.
    ///
    /// The meaning of "point" varies between chart types. For a Line chart a point
    /// is a line segment; in a Column chart a point is a an individual bar; and in
    /// a Pie chart a point is a pie segment.
    ///
    /// A point is represented by the {@link ChartPoint} struct.
    ///
    /// Chart points are most commonly used for Pie and Doughnut charts to format
    /// individual segments of the chart. In all other chart types the formatting
    /// happens at the chart series level.
    ///
    /// # Parameters
    ///
    /// `points`: A slice of {@link ChartPoint} objects.
    #[wasm_bindgen(js_name = "setPoints", skip_jsdoc)]
    pub fn set_points(&self, points: Vec<ChartPoint>) -> ChartSeries {
        let mut lock = self.parent.lock().unwrap();
        lock.add_series()
            .set_points(&points.iter().map(|x| x.inner.clone()).collect::<Vec<_>>());
        ChartSeries {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the colors for points in a chart series.
    ///
    /// As explained above, in the section on {@link ChartSeries#setPoints}, the
    /// most common use case for point formatting is to set the formatting of
    /// individual segments of Pie charts, or in particular to set the colors of
    /// pie segments. For this simple use case the {@link ChartSeries#setPoints}
    /// method can be overly verbose.
    ///
    /// As a syntactic shortcut the `set_point_colors()` method allows you to
    /// set the colors of chart points with a simpler interface.
    ///
    /// Compare the example below with the previous more general example which
    /// both produce the same result.
    ///
    /// # Parameters
    ///
    /// `colors`: a slice of {@link Color} enum values or types that will convert
    /// into {@link Color}.
    #[wasm_bindgen(js_name = "setPointColors", skip_jsdoc)]
    pub fn set_point_colors(&self, colors: Vec<Color>) -> ChartSeries {
        let mut lock = self.parent.lock().unwrap();
        lock.add_series().set_point_colors(
            &colors
                .into_iter()
                .map(|x| xlsx::Color::from(x))
                .collect::<Vec<_>>(),
        );
        ChartSeries {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the trendline for a chart series.
    ///
    /// Excel allows you to add a trendline to a data series that represents the
    /// trend or regression of the data using different types of fit. A
    /// {@link ChartTrendline} struct reference is used to represents the options of
    /// Excel trendlines and can be added to a series via the
    /// {@link ChartSeries#setTrendline} method.
    ///
    /// src="https://rustxlsxwriter.github.io/images/trendline_options.png">
    ///
    /// # Parameters
    ///
    /// `trendline`: A {@link ChartTrendline} reference.
    #[wasm_bindgen(js_name = "setTrendline", skip_jsdoc)]
    pub fn set_trendline(&self, trendline: ChartTrendline) -> ChartSeries {
        let mut lock = self.parent.lock().unwrap();
        lock.add_series()
            .set_trendline(&*trendline.inner.lock().unwrap());
        ChartSeries {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the vertical error bars for a chart series.
    ///
    /// Error bars on Excel charts allow you to show margins of error for a series
    /// based on measures such as Standard Deviation, Standard Error, Fixed values,
    /// Percentages or even custom defined ranges.
    ///
    /// The `ChartErrorBars` struct represents the error bars for a chart series.
    ///
    /// src="https://rustxlsxwriter.github.io/images/chart_error_bars_options.png">
    ///
    /// # Parameters
    ///
    /// `error_bars`: A {@link ChartErrorBars} reference.
    #[wasm_bindgen(js_name = "setYErrorBars", skip_jsdoc)]
    pub fn set_y_error_bars(&self, error_bars: ChartErrorBars) -> ChartSeries {
        let mut lock = self.parent.lock().unwrap();
        lock.add_series()
            .set_y_error_bars(&*error_bars.inner.lock().unwrap());
        ChartSeries {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the horizontal error bars for a chart series.
    ///
    /// See {@link ChartSeries#setYErrorBars} above for a description of error
    /// bars and their properties.
    ///
    /// Horizontal error bars can only be set in Excel for Scatter and Bar charts.
    ///
    /// # Parameters
    ///
    /// `error_bars`: A {@link ChartErrorBars} reference.
    #[wasm_bindgen(js_name = "setXErrorBars", skip_jsdoc)]
    pub fn set_x_error_bars(&self, error_bars: ChartErrorBars) -> ChartSeries {
        let mut lock = self.parent.lock().unwrap();
        lock.add_series()
            .set_x_error_bars(&*error_bars.inner.lock().unwrap());
        ChartSeries {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the series overlap for a chart/bar chart.
    ///
    /// Set the overlap between series in a Bar/Column chart. The range is -100
    /// <= overlap <= 100 and the default is 0.
    ///
    /// Note, In Excel this property is only available for Bar and Column charts
    /// and also only needs to be applied to one of the data series of the
    /// chart.
    ///
    /// # Parameters
    ///
    /// - `overlap`: Overlap percentage of columns in Bar/Column charts. The
    ///   range is -100 <= overlap <= 100 and the default is 0.
    #[wasm_bindgen(js_name = "setOverlap", skip_jsdoc)]
    pub fn set_overlap(&self, overlap: i8) -> ChartSeries {
        let mut lock = self.parent.lock().unwrap();
        lock.add_series().set_overlap(overlap);
        ChartSeries {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the gap width for a chart/bar chart.
    ///
    /// Set the gap width between series in a Bar/Column chart. The range is 0
    /// <= gap <= 500 and the default is 150.
    ///
    /// Note, In Excel this property is only available for Bar and Column charts
    /// and also only needs to be applied to one of the data series of the
    /// chart.
    ///
    /// # Parameters
    ///
    /// - `gap`: Gap percentage of columns in Bar/Column charts. The range is 0
    ///   <= gap <= 500 and the default is 150.
    ///
    /// See the example for {@link ChartSeries#setOverlap} above.
    #[wasm_bindgen(js_name = "setGap", skip_jsdoc)]
    pub fn set_gap(&self, gap: u16) -> ChartSeries {
        let mut lock = self.parent.lock().unwrap();
        lock.add_series().set_gap(gap);
        ChartSeries {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set line type charts to smooth for a series.
    ///
    /// Line and Scatter charts can have a linear or smoothed line connecting
    /// their data points. Some chart types such as {@link ChartType#ScatterSmooth} have
    /// smoothed series by default and some such as
    /// {@link ChartType#ScatterStraight} don't.
    ///
    /// The `ChartSeries::set_smooth()` method can be used to turn on/off the
    /// smooth property for a series.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. The default depends on the chart
    ///   type.
    #[wasm_bindgen(js_name = "setSmooth", skip_jsdoc)]
    pub fn set_smooth(&self, enable: bool) -> ChartSeries {
        let mut lock = self.parent.lock().unwrap();
        lock.add_series().set_smooth(enable);
        ChartSeries {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Invert the color for negative values in a column/bar chart series.
    ///
    /// Bar and Column charts in Excel offer a series property called "Invert if
    /// negative". This isn't available for other types of charts.
    ///
    /// The negative values are shown as a white solid fill with a black border.
    /// To set the color of the negative part of the bar/column see
    /// {@link ChartSeries#setInvertIfNegativeColor} below.
    #[wasm_bindgen(js_name = "setInvertIfNegative", skip_jsdoc)]
    pub fn set_invert_if_negative(&self) -> ChartSeries {
        let mut lock = self.parent.lock().unwrap();
        lock.add_series().set_invert_if_negative();
        ChartSeries {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the inverted color for negative values in a column/bar chart series.
    ///
    /// Bar and Column charts in Excel offer a series property called "Invert if
    /// negative" (see {@link ChartSeries#setInvertIfNegative} above).
    ///
    /// The negative values are usually shown as a white solid fill with a black
    /// border but the `set_invert_if_negative_color()` method can be use to set
    /// a user defined color. This also requires that you set a
    /// {@link ChartSolidFill} for the series.
    ///
    /// # Parameters
    ///
    /// - `color`: The inverse color property defined by a {@link Color} enum value.
    #[wasm_bindgen(js_name = "setInvertIfNegativeColor", skip_jsdoc)]
    pub fn set_invert_if_negative_color(&self, color: Color) -> ChartSeries {
        let mut lock = self.parent.lock().unwrap();
        lock.add_series()
            .set_invert_if_negative_color(xlsx::Color::from(color));
        ChartSeries {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Delete/hide the series name from the chart legend.
    ///
    /// The `delete_from_legend()` method deletes/hides the series name from the
    /// chart legend. This is sometimes required if there are a lot of secondary
    /// series and their names are cluttering the chart legend.
    ///
    /// Note, to hide all the names in the chart legend you should use the
    /// {@link ChartLegend#setHidden} method instead.
    ///
    /// See also the {@link ChartTrendline#deleteFromLegend} and the
    /// {@link ChartLegend#deleteEntries} methods.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "deleteFromLegend", skip_jsdoc)]
    pub fn delete_from_legend(&self, enable: bool) -> ChartSeries {
        let mut lock = self.parent.lock().unwrap();
        lock.add_series().delete_from_legend(enable);
        ChartSeries {
            parent: Arc::clone(&self.parent),
        }
    }
}
