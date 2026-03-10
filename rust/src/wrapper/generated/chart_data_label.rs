use crate::wrapper::ChartFont;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ChartDataLabel` struct represents a chart data label.
///
/// The {@link ChartDataLabel} struct represents the properties of the data labels
/// for a chart series. In Excel a data label can be added to a chart series to
/// indicate the values of the plotted data points. It can also be used to
/// indicate other properties such as the category or, for Pie charts, the
/// percentage.
///
/// It is used in conjunction with the {@link Chart} struct.
///
/// src="https://rustxlsxwriter.github.io/images/chart_data_labels_dialog.png">
#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartDataLabel {
    pub(crate) inner: Arc<Mutex<xlsx::ChartDataLabel>>,
}

#[wasm_bindgen]
impl ChartDataLabel {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ChartDataLabel {
        ChartDataLabel {
            inner: Arc::new(Mutex::new(xlsx::ChartDataLabel::new())),
        }
    }
    /// Display the point value on the data label.
    ///
    /// This method turns on the option to display the value of the data point.
    ///
    /// If no other display options is set, such as `show_category()`, then this
    /// value will default to on, like in Excel.
    #[wasm_bindgen(js_name = "showValue", skip_jsdoc)]
    pub fn show_value(&self) -> ChartDataLabel {
        let mut lock = self.inner.lock().unwrap();
        lock.show_value();
        ChartDataLabel {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Display the point category name on the data label.
    ///
    /// This method turns on the option to display the category name of the data
    /// point.
    #[wasm_bindgen(js_name = "showCategoryName", skip_jsdoc)]
    pub fn show_category_name(&self) -> ChartDataLabel {
        let mut lock = self.inner.lock().unwrap();
        lock.show_category_name();
        ChartDataLabel {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Display the series name on the data label.
    #[wasm_bindgen(js_name = "showSeriesName", skip_jsdoc)]
    pub fn show_series_name(&self) -> ChartDataLabel {
        let mut lock = self.inner.lock().unwrap();
        lock.show_series_name();
        ChartDataLabel {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Display leader lines from/to the data label.
    ///
    /// **Note**: Even when leader lines are turned on they are not
    /// automatically visible in a chart. In Excel a leader line only appears if
    /// the data label is moved manually or if the data labels are very close
    /// and need to be adjusted automatically.
    #[wasm_bindgen(js_name = "showLeaderLines", skip_jsdoc)]
    pub fn show_leader_lines(&self) -> ChartDataLabel {
        let mut lock = self.inner.lock().unwrap();
        lock.show_leader_lines();
        ChartDataLabel {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Show the legend key/symbol on the data label.
    #[wasm_bindgen(js_name = "showLegendKey", skip_jsdoc)]
    pub fn show_legend_key(&self) -> ChartDataLabel {
        let mut lock = self.inner.lock().unwrap();
        lock.show_legend_key();
        ChartDataLabel {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Display the chart value as a percentage.
    ///
    /// This method is used to turn on the display of data labels as a
    /// percentage for a series. It is mainly used for pie charts.
    #[wasm_bindgen(js_name = "showPercentage", skip_jsdoc)]
    pub fn show_percentage(&self) -> ChartDataLabel {
        let mut lock = self.inner.lock().unwrap();
        lock.show_percentage();
        ChartDataLabel {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the font properties of a chart data label.
    ///
    /// Set the font properties of a chart data labels using a {@link ChartFont}
    /// reference. Example font properties that can be set are:
    ///
    /// - {@link ChartFont#setBold}
    /// - {@link ChartFont#setItalic}
    /// - {@link ChartFont#setName}
    /// - {@link ChartFont#setSize}
    /// - {@link ChartFont#setRotation}
    ///
    /// See {@link ChartFont} for full details. ///
    /// # Parameters
    ///
    /// `font`: A {@link ChartFont} struct reference to represent the font
    /// properties.
    #[wasm_bindgen(js_name = "setFont", skip_jsdoc)]
    pub fn set_font(&self, font: ChartFont) -> ChartDataLabel {
        let mut lock = self.inner.lock().unwrap();
        lock.set_font(&font.inner);
        ChartDataLabel {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the number format for a chart data label.
    ///
    /// Excel plots/displays data in charts with the same number formatting that
    /// it has in the worksheet. The `set_num_format()` method allows you to
    /// override this and controls whether a number is displayed as an integer,
    /// a floating point number, a date, a currency value or some other user
    /// defined format.
    ///
    /// See also [Number Format Categories] and [Number Formats in different
    /// locales] in the documentation for {@link Format}.
    ///
    /// [Number Format Categories]: crate::Format#number-format-categories
    /// [Number Formats in different locales]:
    ///     crate::Format#number-formats-in-different-locales
    ///
    /// # Parameters
    ///
    /// - `num_format`: The number format property.
    #[wasm_bindgen(js_name = "setNumFormat", skip_jsdoc)]
    pub fn set_num_format(&self, num_format: &str) -> ChartDataLabel {
        let mut lock = self.inner.lock().unwrap();
        lock.set_num_format(num_format);
        ChartDataLabel {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Display the point Y value on the data label.
    ///
    /// This is the same as {@link ChartDataLabel#showValue} except it is named
    /// differently in Excel for Scatter charts. The methods are equivalent and
    /// either one can be used for any chart type.
    #[wasm_bindgen(js_name = "showYValue", skip_jsdoc)]
    pub fn show_y_value(&self) -> ChartDataLabel {
        let mut lock = self.inner.lock().unwrap();
        lock.show_y_value();
        ChartDataLabel {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Display the point X value on the data label.
    ///
    /// This is the same as {@link ChartDataLabel#showCategoryName} except it
    /// is named differently in Excel for Scatter charts. The methods are
    /// equivalent and either one can be used for any chart type.
    #[wasm_bindgen(js_name = "showXValue", skip_jsdoc)]
    pub fn show_x_value(&self) -> ChartDataLabel {
        let mut lock = self.inner.lock().unwrap();
        lock.show_x_value();
        ChartDataLabel {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set a custom data label as hidden.
    ///
    /// This method hides a custom data label used with the
    /// {@link ChartSeries#setCustomDataLabels} method. It is ignored if used
    /// with a series {@link ChartDataLabel}.
    #[wasm_bindgen(js_name = "setHidden", skip_jsdoc)]
    pub fn set_hidden(&self) -> ChartDataLabel {
        let mut lock = self.inner.lock().unwrap();
        lock.set_hidden();
        ChartDataLabel {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Turn a data label reference into a custom data label.
    ///
    /// Converts a `&ChartDataLabel` reference into a {@link ChartDataLabel} so that
    /// it can be used as a custom data label with the
    /// {@link ChartSeries#setCustomDataLabels} method.
    ///
    /// This is a syntactic shortcut for a simple `clone()`.
    #[wasm_bindgen(js_name = "toCustom", skip_jsdoc)]
    pub fn to_custom(&self) -> ChartDataLabel {
        let mut lock = self.inner.lock().unwrap();
        lock.to_custom();
        ChartDataLabel {
            inner: Arc::clone(&self.inner),
        }
    }
}
