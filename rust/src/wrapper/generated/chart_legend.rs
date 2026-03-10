use crate::wrapper::ChartFont;
use crate::wrapper::ChartLayout;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ChartLegend` struct represents a chart legend.
///
/// The `ChartLegend` struct is a representation of a legend on an Excel chart.
/// The legend is a rectangular box that identifies the name and color of each
/// of the series in the chart.
///
/// `ChartLegend` can be used to configure properties of the chart legend and is
/// usually obtained via the {@link Chart#legend} method.
///
/// It is used in conjunction with the {@link Chart} struct.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartLegend {
    pub(crate) parent: Arc<Mutex<xlsx::Chart>>,
}

#[wasm_bindgen]
impl ChartLegend {
    /// Hide the legend for a Chart.
    ///
    /// The legend if usually on by default for an Excel chart. The
    /// `set_hidden()` method can be used to turn it off.
    #[wasm_bindgen(js_name = "setHidden", skip_jsdoc)]
    pub fn set_hidden(&self) -> ChartLegend {
        let mut lock = self.parent.lock().unwrap();
        lock.legend().set_hidden();
        ChartLegend {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the chart legend as overlaid on the chart.
    ///
    /// In the Excel "Format Legend" dialog there is a default option of "Show
    /// the legend without overlapping the chart":
    ///
    /// This can be turned off using the `set_overlay()` method.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setOverlay", skip_jsdoc)]
    pub fn set_overlay(&self, enable: bool) -> ChartLegend {
        let mut lock = self.parent.lock().unwrap();
        lock.legend().set_overlay(enable);
        ChartLegend {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the manual position of the chart axis legend.
    ///
    /// This method is used to simulate manual positioning of a chart legend.
    /// See {@link ChartLayout} for more details.
    ///
    /// Note, to position the title over the plot area of the chart you will
    /// also need to set the {@link ChartLegend#setOverlay} property.
    ///
    /// # Parameters
    ///
    /// - `layout`: A {@link ChartLayout} struct reference.
    #[wasm_bindgen(js_name = "setLayout", skip_jsdoc)]
    pub fn set_layout(&self, layout: ChartLayout) -> ChartLegend {
        let mut lock = self.parent.lock().unwrap();
        lock.legend().set_layout(&layout.inner);
        ChartLegend {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the font properties of a chart legend.
    ///
    /// Set the font properties of a chart legend using a {@link ChartFont}
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
    /// `font` - A {@link ChartFont} struct reference to represent the font
    /// properties.
    #[wasm_bindgen(js_name = "setFont", skip_jsdoc)]
    pub fn set_font(&self, font: ChartFont) -> ChartLegend {
        let mut lock = self.parent.lock().unwrap();
        lock.legend().set_font(&font.inner);
        ChartLegend {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Delete/hide series names from the chart legend.
    ///
    /// The `delete_entries()` method deletes/hides one or more series names
    /// from the chart legend. This is sometimes required if there are a lot of
    /// secondary series and their names are cluttering the chart legend.
    ///
    /// The same effect can be accomplished using the
    /// {@link ChartSeries#deleteFromLegend} and
    /// {@link ChartTrendline#deleteFromLegend} methods. However, this method
    /// can be used for some edge cases such as Pie/Doughnut charts which
    /// display legend entries for each point in the series.
    ///
    /// Note, to hide all the names in the chart legend you should use the
    /// {@link ChartLegend#setHidden} method instead.
    ///
    /// # Parameters
    ///
    /// - `entries`: A slice ref of `usize` zero-indexed indices of the
    ///   series names to be hidden.
    #[wasm_bindgen(js_name = "deleteEntries", skip_jsdoc)]
    pub fn delete_entries(&self, entries: Vec<usize>) -> ChartLegend {
        let mut lock = self.parent.lock().unwrap();
        lock.legend().delete_entries(&entries);
        ChartLegend {
            parent: Arc::clone(&self.parent),
        }
    }
}
