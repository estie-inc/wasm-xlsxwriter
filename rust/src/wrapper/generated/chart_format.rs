use crate::wrapper::ChartGradientFill;
use crate::wrapper::ChartLine;
use crate::wrapper::ChartPatternFill;
use crate::wrapper::ChartSolidFill;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ChartFormat` struct represents formatting for various chart objects.
///
/// Excel uses a standard formatting dialog for the elements of a chart such as
/// data series, the plot area, the chart area, the legend or individual points.
/// It looks like this:
///
/// The {@link ChartFormat} struct represents many of these format options and just
/// like Excel it offers a similar formatting interface for a number of the
/// chart sub-elements supported by `rust_xlsxwriter`.
///
/// It is used in conjunction with the {@link Chart} struct.
///
/// The {@link ChartFormat} struct is generally passed to the `set_format()` method
/// of a chart element. It supports several chart formatting elements such as:
///
/// - {@link ChartFormat#setSolidFill}: Set the {@link ChartSolidFill} properties.
/// - {@link ChartFormat#setPatternFill}: Set the {@link ChartPatternFill}
///   properties.
/// - {@link ChartFormat#setGradientFill}: Set the {@link ChartGradientFill}
///   properties.
/// - {@link ChartFormat#setNoFill}: Turn off the fill for the chart object.
/// - {@link ChartFormat#setLine}: Set the {@link ChartLine} properties.
/// - {@link ChartFormat#setBorder}: Set the {@link ChartBorder} properties. A
///   synonym for {@link ChartLine} depending on context.
/// - {@link ChartFormat#setNoLine}: Turn off the line for the chart object.
/// - {@link ChartFormat#setNoBorder}: Turn off the border for the chart
///   object.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartFormat {
    pub(crate) inner: Arc<Mutex<xlsx::ChartFormat>>,
}

#[wasm_bindgen]
impl ChartFormat {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ChartFormat {
        ChartFormat {
            inner: Arc::new(Mutex::new(xlsx::ChartFormat::new())),
        }
    }
    #[doc = r" Create a deep clone of this object."]
    #[wasm_bindgen(js_name = "clone")]
    pub fn deep_clone(&self) -> ChartFormat {
        ChartFormat {
            inner: Arc::new(Mutex::new(self.inner.lock().unwrap().clone())),
        }
    }
    /// Set the line formatting for a chart element.
    ///
    /// See the {@link ChartLine} struct for details on the line properties that can
    /// be set.
    ///
    /// # Parameters
    ///
    /// - `line`: A {@link ChartLine} struct reference.
    #[wasm_bindgen(js_name = "setLine", skip_jsdoc)]
    pub fn set_line(&self, line: &ChartLine) -> ChartFormat {
        let mut lock = self.inner.lock().unwrap();
        lock.set_line(&*line.inner.lock().unwrap());
        ChartFormat {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the border formatting for a chart element.
    ///
    /// See the {@link ChartLine} struct for details on the border properties that
    /// can be set. As a syntactic shortcut you can use the type alias
    /// {@link ChartBorder} instead
    /// of `ChartLine`.
    ///
    /// # Parameters
    ///
    /// - `line`: A {@link ChartLine} struct reference.
    #[wasm_bindgen(js_name = "setBorder", skip_jsdoc)]
    pub fn set_border(&self, line: &ChartLine) -> ChartFormat {
        let mut lock = self.inner.lock().unwrap();
        lock.set_border(&*line.inner.lock().unwrap());
        ChartFormat {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Turn off the line property for a chart element.
    ///
    /// The line property for a chart element can be turned off if you wish to
    /// hide it.
    #[wasm_bindgen(js_name = "setNoLine", skip_jsdoc)]
    pub fn set_no_line(&self) -> ChartFormat {
        let mut lock = self.inner.lock().unwrap();
        lock.set_no_line();
        ChartFormat {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Turn off the border property for a chart element.
    ///
    /// The border property for a chart element can be turned off if you wish to
    /// hide it.
    #[wasm_bindgen(js_name = "setNoBorder", skip_jsdoc)]
    pub fn set_no_border(&self) -> ChartFormat {
        let mut lock = self.inner.lock().unwrap();
        lock.set_no_border();
        ChartFormat {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Turn off the fill property for a chart element.
    ///
    /// The fill property for a chart element can be turned off if you wish to
    /// hide it and display only the border (if set).
    #[wasm_bindgen(js_name = "setNoFill", skip_jsdoc)]
    pub fn set_no_fill(&self) -> ChartFormat {
        let mut lock = self.inner.lock().unwrap();
        lock.set_no_fill();
        ChartFormat {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the solid fill formatting for a chart element.
    ///
    /// See the {@link ChartSolidFill} struct for details on the solid fill
    /// properties that can be set.
    ///
    /// # Parameters
    ///
    /// - `fill`: A {@link ChartSolidFill} struct reference.
    #[wasm_bindgen(js_name = "setSolidFill", skip_jsdoc)]
    pub fn set_solid_fill(&self, fill: &ChartSolidFill) -> ChartFormat {
        let mut lock = self.inner.lock().unwrap();
        lock.set_solid_fill(&*fill.inner.lock().unwrap());
        ChartFormat {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the pattern fill formatting for a chart element.
    ///
    /// See the {@link ChartPatternFill} struct for details on the pattern fill
    /// properties that can be set.
    ///
    /// # Parameters
    ///
    /// - `fill`: A {@link ChartPatternFill} struct reference.
    #[wasm_bindgen(js_name = "setPatternFill", skip_jsdoc)]
    pub fn set_pattern_fill(&self, fill: &ChartPatternFill) -> ChartFormat {
        let mut lock = self.inner.lock().unwrap();
        lock.set_pattern_fill(&*fill.inner.lock().unwrap());
        ChartFormat {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the gradient fill formatting for a chart element.
    ///
    /// See the {@link ChartGradientFill} struct for details on the gradient fill
    /// properties that can be set.
    ///
    /// # Parameters
    ///
    /// - `fill`: A {@link ChartGradientFill} struct reference.
    #[wasm_bindgen(js_name = "setGradientFill", skip_jsdoc)]
    pub fn set_gradient_fill(&self, fill: &ChartGradientFill) -> ChartFormat {
        let mut lock = self.inner.lock().unwrap();
        lock.set_gradient_fill(&*fill.inner.lock().unwrap());
        ChartFormat {
            inner: Arc::clone(&self.inner),
        }
    }
}
