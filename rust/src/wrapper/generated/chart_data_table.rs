use crate::wrapper::ChartFont;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ChartDataTable` struct represents an optional data table displayed
/// below the chart.
///
/// A chart data table in Excel is an additional table below a chart that shows
/// the plotted data in tabular form.
///
/// The chart data table has the following default properties which can be set
/// with the methods outlined below.
///
/// The `ChartDataTable` struct is used in conjunction with the
/// {@link Chart#setDataTable} method.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartDataTable {
    pub(crate) inner: Arc<Mutex<xlsx::ChartDataTable>>,
}

#[wasm_bindgen]
impl ChartDataTable {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ChartDataTable {
        ChartDataTable {
            inner: Arc::new(Mutex::new(xlsx::ChartDataTable::new())),
        }
    }
    /// Turn on/off the horizontal border lines for a chart data table.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is on by default.
    #[wasm_bindgen(js_name = "showHorizontalBorders", skip_jsdoc)]
    pub fn show_horizontal_borders(&self, enable: bool) -> ChartDataTable {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.show_horizontal_borders(enable);
        *lock = inner;
        ChartDataTable {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Turn on/off the vertical border lines for a chart data table.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is on by default.
    #[wasm_bindgen(js_name = "showVerticalBorders", skip_jsdoc)]
    pub fn show_vertical_borders(&self, enable: bool) -> ChartDataTable {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.show_vertical_borders(enable);
        *lock = inner;
        ChartDataTable {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Turn on/off the outline border lines for a chart data table.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is on by default.
    #[wasm_bindgen(js_name = "showOutlineBorders", skip_jsdoc)]
    pub fn show_outline_borders(&self, enable: bool) -> ChartDataTable {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.show_outline_borders(enable);
        *lock = inner;
        ChartDataTable {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Turn on/off the legend keys for a chart data table.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "showLegendKeys", skip_jsdoc)]
    pub fn show_legend_keys(&self, enable: bool) -> ChartDataTable {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.show_legend_keys(enable);
        *lock = inner;
        ChartDataTable {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the font properties of a chart data table.
    ///
    /// Set the font properties of a chart data table using a {@link ChartFont}
    /// reference. Example font properties that can be set are:
    ///
    /// - {@link ChartFont#setBold}
    /// - {@link ChartFont#setItalic}
    /// - {@link ChartFont#setColor}
    /// - {@link ChartFont#setName}
    /// - {@link ChartFont#setSize}
    /// - {@link ChartFont#setRotation}
    /// - {@link ChartFont#setUnderline}
    /// - {@link ChartFont#setStrikethrough}
    /// - {@link ChartFont#setRightToLeft}
    ///
    /// See {@link ChartFont} for full details.
    ///
    /// # Parameters
    ///
    /// `font`: A {@link ChartFont} struct reference to represent the font
    /// properties.
    #[wasm_bindgen(js_name = "setFont", skip_jsdoc)]
    pub fn set_font(&self, font: ChartFont) -> ChartDataTable {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_font(&font.inner);
        *lock = inner;
        ChartDataTable {
            inner: Arc::clone(&self.inner),
        }
    }
}
