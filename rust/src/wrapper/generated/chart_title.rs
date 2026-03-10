use crate::wrapper::ChartFont;
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
    pub fn set_font(&self, font: ChartFont) -> ChartTitle {
        let mut lock = self.parent.lock().unwrap();
        lock.title().set_font(&font.inner);
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
    pub fn set_layout(&self, layout: ChartLayout) -> ChartTitle {
        let mut lock = self.parent.lock().unwrap();
        lock.title().set_layout(&layout.inner);
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
