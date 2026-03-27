use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ChartLayout` struct represents the dimensions required to position some
/// chart objects.
///
/// Excel allows the user to positions chart objects like axis labels or the
/// legend is two ways.
///
/// The first method is via standard positions such as top, bottom, left and
/// right. The `rust_xlsxwriter` library replicates this via enums such as
/// {@link ChartAxisLabelPosition} and {@link ChartLegendPosition} and the associated
/// methods that use them.
///
/// The second method Excel supports is manual positioning of elements such as
/// the chart axis labels, the chart legend, the chart plot area and the chart
/// title. The `rust_xlsxwriter` library replicates this type of positioning via
/// the `ChartLayout` struct.
///
/// The layout units used by Excel are relative units expressed as a percentage
/// of the chart dimensions and are `f64` values in the range `0.0 < x <= 1.0`.
/// Excel calculates these dimensions as shown below:
///
/// With reference to the above figure the layout units are calculated as
/// follows:
///
/// These units are cumbersome and can vary depending on other elements in the
/// chart such as text lengths. However, these are the units that are required
/// by Excel to allow relative positioning. Some trial and error is generally
/// required.
///
/// For {@link ChartPlotArea} and {@link ChartLegend} you can also set the width and
/// height based on the following calculation:
///
/// For other text based objects the width and height are changed by the font
/// dimensions.
///
/// The {@link Chart} objects and methods that support `ChartLayout` are:
///
/// - {@link ChartTitle#setLayout}
/// - {@link ChartLegend#setLayout}
/// - {@link ChartPlotArea#setLayout}
/// - {@link ChartAxis#setLabelLayout}
#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartLayout {
    pub(crate) inner: Arc<Mutex<xlsx::ChartLayout>>,
}

#[wasm_bindgen]
impl ChartLayout {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ChartLayout {
        ChartLayout {
            inner: Arc::new(Mutex::new(xlsx::ChartLayout::new())),
        }
    }
    #[doc = r" Create a deep clone of this object."]
    #[wasm_bindgen(js_name = "clone")]
    pub fn deep_clone(&self) -> ChartLayout {
        ChartLayout {
            inner: Arc::new(Mutex::new(self.inner.lock().unwrap().clone())),
        }
    }
    /// Set the offset for a layout.
    ///
    /// With reference to the figure below the layout units are calculated as
    /// follows:
    ///
    /// See {@link ChartLayout} above for an explanation of chart layouts and the
    /// units used.
    ///
    /// # Parameters
    ///
    /// - `x_offset`: A `f64` value in the range `0.0 < x <= 1.0`.
    /// - `y_offset`: A `f64` value in the range `0.0 < y <= 1.0`.
    #[wasm_bindgen(js_name = "setOffset", skip_jsdoc)]
    pub fn set_offset(&self, x_offset: f64, y_offset: f64) -> ChartLayout {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_offset(x_offset, y_offset);
        *lock = inner;
        ChartLayout {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the dimensions for a layout.
    ///
    /// With reference to the figure above the dimension units are calculated as
    /// follows:
    ///
    /// The dimensions are only used for {@link ChartPlotArea} and {@link ChartLegend}.
    /// For other text based objects the width and height are changed by the
    /// font dimensions.
    ///
    /// See {@link ChartLayout} above for an explanation of chart layouts and the
    /// units used.
    ///
    /// # Parameters
    ///
    /// - `width`: A `f64` value in the range `0.0 < width <= 1.0`.
    /// - `height`: A `f64` value in the range `0.0 < height <= 1.0`.
    #[wasm_bindgen(js_name = "setDimensions", skip_jsdoc)]
    pub fn set_dimensions(&self, width: f64, height: f64) -> ChartLayout {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_dimensions(width, height);
        *lock = inner;
        ChartLayout {
            inner: Arc::clone(&self.inner),
        }
    }
}
