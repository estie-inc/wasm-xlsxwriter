use crate::wrapper::ChartFont;
use crate::wrapper::ChartFormat;
use crate::wrapper::ChartLayout;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ChartTitle` struct represents a chart title.
///
/// It is used in conjunction with the {@link Chart} struct.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartTitle {
    pub(crate) parent: Arc<Mutex<xlsx::Chart>>,
}

#[wasm_bindgen]
impl ChartTitle {
    /// Add a title for a chart.
    ///
    /// Set the name (title) for the chart. The name is displayed above the
    /// chart.
    ///
    /// The name can be a simple string, a formula such as `Sheet1!$A$1` or a
    /// tuple with a sheet name, row and column such as `('Sheet1', 0, 0)`.
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
    #[wasm_bindgen(js_name = "setName", skip_jsdoc)]
    pub fn set_name(&self, name: &str) -> ChartTitle {
        let mut lock = self.parent.lock().unwrap();
        lock.title().set_name(name);
        ChartTitle {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Hide the chart title.
    ///
    /// By default Excel adds an automatic chart title to charts with a single
    /// series and a user defined series name. The `set_hidden()` method turns
    /// this default title off.
    #[wasm_bindgen(js_name = "setHidden", skip_jsdoc)]
    pub fn set_hidden(&self) -> ChartTitle {
        let mut lock = self.parent.lock().unwrap();
        lock.title().set_hidden();
        ChartTitle {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the formatting properties for a chart title.
    ///
    /// Set the formatting properties for a chart title via a {@link ChartFormat}
    /// object or a sub struct that implements {@link IntoChartFormat}.
    ///
    /// The formatting that can be applied via a {@link ChartFormat} object are:
    ///
    /// - {@link ChartFormat#setSolidFill}: Set the {@link ChartSolidFill} properties.
    /// - {@link ChartFormat#setPatternFill}: Set the {@link ChartPatternFill} properties.
    /// - {@link ChartFormat#setGradientFill}: Set the {@link ChartGradientFill} properties.
    /// - {@link ChartFormat#setNoFill}: Turn off the fill for the chart object.
    /// - {@link ChartFormat#setLine}: Set the {@link ChartLine} properties.
    /// - {@link ChartFormat#setBorder}: Set the {@link ChartBorder} properties.
    ///   A synonym for {@link ChartLine} depending on context.
    /// - {@link ChartFormat#setNoLine}: Turn off the line for the chart object.
    /// - {@link ChartFormat#setNoBorder}: Turn off the border for the chart object.
    ///
    /// # Parameters
    ///
    /// `format`: A {@link ChartFormat} struct reference or a sub struct that will
    /// convert into a `ChartFormat` instance. See the docs for
    /// {@link IntoChartFormat} for details.
    #[wasm_bindgen(js_name = "setFormat", skip_jsdoc)]
    pub fn set_format(&self, format: &ChartFormat) -> ChartTitle {
        let mut lock = self.parent.lock().unwrap();
        lock.title().set_format(&mut *format.inner.lock().unwrap());
        ChartTitle {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the font properties of a chart title.
    ///
    /// Set the font properties of a chart title using a {@link ChartFont}
    /// reference. Example font properties that can be set are:
    ///
    /// - {@link ChartFont#setBold}
    /// - {@link ChartFont#setItalic}
    /// - {@link ChartFont#setName}
    /// - {@link ChartFont#setSize}
    /// - {@link ChartFont#setRotation}
    ///
    /// See {@link ChartFont} for full details.
    ///
    /// # Parameters
    ///
    /// `font`: A {@link ChartFont} struct reference to represent the font
    /// properties.
    #[wasm_bindgen(js_name = "setFont", skip_jsdoc)]
    pub fn set_font(&self, font: &ChartFont) -> ChartTitle {
        let mut lock = self.parent.lock().unwrap();
        lock.title().set_font(&*font.inner.lock().unwrap());
        ChartTitle {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the manual position of the chart title.
    ///
    /// This method is used to simulate manual positioning of a chart title. See
    /// {@link ChartLayout} for more details.
    ///
    /// Note, to position the title over the plot area of the chart you will
    /// also need to set the {@link ChartTitle#setOverlay} property.
    ///
    /// # Parameters
    ///
    /// - `layout`: A {@link ChartLayout} struct reference.
    #[wasm_bindgen(js_name = "setLayout", skip_jsdoc)]
    pub fn set_layout(&self, layout: &ChartLayout) -> ChartTitle {
        let mut lock = self.parent.lock().unwrap();
        lock.title().set_layout(&*layout.inner.lock().unwrap());
        ChartTitle {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the chart title as overlaid on the chart.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setOverlay", skip_jsdoc)]
    pub fn set_overlay(&self, enable: bool) -> ChartTitle {
        let mut lock = self.parent.lock().unwrap();
        lock.title().set_overlay(enable);
        ChartTitle {
            parent: Arc::clone(&self.parent),
        }
    }
}
