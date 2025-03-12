use std::sync::{Arc, Mutex};

use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

use super::{chart_data_label::ChartDataLabel, chart_marker::ChartMarker, chart_point::ChartPoint, chart_range::ChartRange};

#[wasm_bindgen]
pub struct ChartSeries {
    pub(crate) inner: Arc<Mutex<xlsx::ChartSeries>>,
}

#[wasm_bindgen]
impl ChartSeries {
    /// Create a new chart series object.
    ///
    /// Create a new chart series object. A chart in Excel must contain at least
    /// one data series. The `ChartSeries` struct represents the category and
    /// value ranges, and the formatting and options for the chart series.
    ///
    /// It is used in conjunction with the {@link Chart} struct.
    ///
    /// A chart series is usually created via the {@link Chart#addSeries}
    /// method, see the first example below. However, if required you can create
    /// a standalone `ChartSeries` object and add it to a chart via the
    /// {@link Chart#pushSeries} method, see the second example below.
    ///
    /// TODO: Add examples
    #[wasm_bindgen(constructor)]
    pub fn new() -> ChartSeries {
        ChartSeries {
            inner: Arc::new(Mutex::new(xlsx::ChartSeries::new())),
        }
    }

    /// Add a category range chart series.
    ///
    /// This method sets the chart category labels. The category is more or less
    /// the same as the X axis. In most chart types the categories property is
    /// optional and the chart will just assume a sequential series from `1..n`.
    /// The exception to this is the Scatter chart types for which a category
    /// range is mandatory in Excel.
    ///
    /// The data range can be set using a formula as shown in the first part of
    /// the example below or using a list of values as shown in the second part.
    ///
    /// # Parameters
    ///
    /// - `range`: The range property which can be one of two generic types:
    ///    - A string with an Excel like range formula such as
    ///      `"Sheet1!$A$1:$A$3"`.
    ///    - A tuple that can be used to create the range programmatically using
    ///      a sheet name and zero indexed row and column values like:
    ///      `("Sheet1", 0, 0, 2, 0)` (this gives the same range as the previous
    ///      string value).
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "setCategories", skip_jsdoc)]
    pub fn set_categories(&self, range: &ChartRange) -> ChartSeries {
        let mut series = self.inner.lock().unwrap();
        series.set_categories(&range.inner);
        ChartSeries {
            inner: Arc::clone(&self.inner),
        }
    }

    /// Add a name for a chart series.
    ///
    /// Set the name for the series. The name is displayed in the formula bar.
    /// For non-Pie/Doughnut charts it is also displayed in the legend. The name
    /// property is optional and if it isn’t supplied it will default to `Series
    /// 1..n`. The name can be a simple string, a formula such as `Sheet1!$A$1`
    /// or a tuple with a sheet name, row and column such as `('Sheet1', 0, 0)`.
    ///
    /// # Parameters
    ///
    /// - `range`: The range property which can be one of the following generic
    ///   types:
    ///    - A simple string title.
    ///    - A string with an Excel like range formula such as `"Sheet1!$A$1"`.
    ///    - A tuple that can be used to create the range programmatically using
    ///      a sheet name and zero indexed row and column values like:
    ///      `("Sheet1", 0, 0)` (this gives the same range as the previous
    ///      string value).
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "setName", skip_jsdoc)]
    pub fn set_name(&self, name: &str) -> ChartSeries {
        let mut series = self.inner.lock().unwrap();
        series.set_name(name);
        ChartSeries {
            inner: Arc::clone(&self.inner),
        }
    }

    /// Add a values range to a chart series.
    ///
    /// All chart series in Excel must have a data range that defines the range
    /// of values for the series. In Excel this is typically a range like
    /// `"Sheet1!$B$1:$B$5"`.
    ///
    /// This is the most important property of a series and is the only
    /// mandatory option for every chart object. This series values links the
    /// chart with the worksheet data that it displays. The data range can be
    /// set using a formula as shown in the first part of the example below or
    /// using a list of values as shown in the second part.
    ///
    /// # Parameters
    ///
    /// - `range`: The range property which can be one of two generic types:
    ///    - A string with an Excel like range formula such as
    ///      `"Sheet1!$A$1:$A$3"`.
    ///    - A tuple that can be used to create the range programmatically using
    ///      a sheet name and zero indexed row and column values like:
    ///      `("Sheet1", 0, 0, 2, 0)` (this gives the same range as the previous
    ///      string value).
    #[wasm_bindgen(js_name = "setValues", skip_jsdoc)]
    pub fn set_values(&self, range: &ChartRange) -> ChartSeries {
        let mut series = self.inner.lock().unwrap();
        series.set_values(&range.inner);
        ChartSeries {
            inner: Arc::clone(&self.inner),
        }
    }

    // FIXME: ownership?
    #[wasm_bindgen(js_name = "setPoints", skip_jsdoc)]
    pub fn set_points(&self, points: Vec<ChartPoint>) -> ChartSeries {
        let mut series = self.inner.lock().unwrap();
        let points: Vec<_> = points.iter().map(|p| p.inner.clone()).collect();
        series.set_points(&points);
        ChartSeries {
            inner: Arc::clone(&self.inner),
        }
    }

    #[wasm_bindgen(js_name = "setDataLabel", skip_jsdoc)]
    pub fn set_data_label(&self, data_label: &ChartDataLabel) -> ChartSeries {
        let mut series = self.inner.lock().unwrap();
        series.set_data_label(&data_label.inner);
        ChartSeries {
            inner: Arc::clone(&self.inner),
        }
    }

    #[wasm_bindgen(js_name = "setMarker", skip_jsdoc)]
    pub fn set_marker(&self, marker: &ChartMarker) -> ChartSeries {
        let mut series = self.inner.lock().unwrap();
        series.set_marker(&marker.inner);
        ChartSeries {
            inner: Arc::clone(&self.inner),
        }
    }
}
