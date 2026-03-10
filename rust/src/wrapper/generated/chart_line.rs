use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ChartLine` struct represents a chart line/border.
///
/// The {@link ChartLine} struct represents the formatting properties for a line or
/// border for a Chart element. It is a sub property of the {@link ChartFormat}
/// struct and is used with the {@link ChartFormat#setLine} or
/// {@link ChartFormat#setBorder} methods.
///
/// Excel uses the element names "Line" and "Border" depending on the context.
/// For a Line chart the line is represented by a line property but for a Column
/// chart the line becomes the border. Both of these share the same properties
/// and are both represented in `rust_xlsxwriter` by the {@link ChartLine} struct.
///
/// As a syntactic shortcut you can use the type alias {@link ChartBorder} instead
/// of `ChartLine`.
///
/// It is used in conjunction with the {@link Chart} struct.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartLine {
    pub(crate) inner: Arc<Mutex<xlsx::ChartLine>>,
}

#[wasm_bindgen]
impl ChartLine {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ChartLine {
        ChartLine {
            inner: Arc::new(Mutex::new(xlsx::ChartLine::new())),
        }
    }
    /// Set the width of the line or border.
    ///
    /// # Parameters
    ///
    /// - `width`: The width should be specified in increments of 0.25 of a
    ///   point as in Excel. The width can be an number type that convert
    ///   `Into` `f64`.
    #[wasm_bindgen(js_name = "setWidth", skip_jsdoc)]
    pub fn set_width(&self, width: f64) -> ChartLine {
        let mut lock = self.inner.lock().unwrap();
        lock.set_width(width);
        ChartLine {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the transparency of a line/border.
    ///
    /// Set the transparency of a line/border for a Chart element. You must also
    /// specify a line color in order for the transparency to be applied.
    ///
    /// # Parameters
    ///
    /// - `transparency`: The color transparency in the range 0 <= transparency
    ///   <= 100. The default value is 0.
    #[wasm_bindgen(js_name = "setTransparency", skip_jsdoc)]
    pub fn set_transparency(&self, transparency: u8) -> ChartLine {
        let mut lock = self.inner.lock().unwrap();
        lock.set_transparency(transparency);
        ChartLine {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the chart line as hidden.
    ///
    /// The method is sometimes required to turn off a default line type in
    /// order to highlight some other element such as the line markers.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off (not hidden) by
    ///   default.
    #[wasm_bindgen(js_name = "setHidden", skip_jsdoc)]
    pub fn set_hidden(&self, enable: bool) -> ChartLine {
        let mut lock = self.inner.lock().unwrap();
        lock.set_hidden(enable);
        ChartLine {
            inner: Arc::clone(&self.inner),
        }
    }
}
