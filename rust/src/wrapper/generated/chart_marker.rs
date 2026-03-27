use crate::wrapper::ChartFormat;
use crate::wrapper::ChartMarkerType;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ChartMarker` struct represents a chart marker.
///
/// The {@link ChartMarker} struct represents the properties of a marker on a Line,
/// Scatter or Radar chart. In Excel a marker is a shape that represents a data
/// point in a chart series.
///
/// It is used in conjunction with the {@link Chart} struct.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartMarker {
    pub(crate) inner: Arc<Mutex<xlsx::ChartMarker>>,
}

#[wasm_bindgen]
impl ChartMarker {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ChartMarker {
        ChartMarker {
            inner: Arc::new(Mutex::new(xlsx::ChartMarker::new())),
        }
    }
    #[doc = r" Create a deep clone of this object."]
    #[wasm_bindgen(js_name = "clone")]
    pub fn deep_clone(&self) -> ChartMarker {
        ChartMarker {
            inner: Arc::new(Mutex::new(self.inner.lock().unwrap().clone())),
        }
    }
    /// Set the automatic/default marker type.
    ///
    /// Allow the marker type to be set automatically by Excel.
    #[wasm_bindgen(js_name = "setAutomatic", skip_jsdoc)]
    pub fn set_automatic(&self) -> ChartMarker {
        let mut lock = self.inner.lock().unwrap();
        lock.set_automatic();
        ChartMarker {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Turn off/hide a chart marker.
    ///
    /// This method can be use to turn off markers for an individual data series
    /// in a chart that has default markers for all series.
    #[wasm_bindgen(js_name = "setNone", skip_jsdoc)]
    pub fn set_none(&self) -> ChartMarker {
        let mut lock = self.inner.lock().unwrap();
        lock.set_none();
        ChartMarker {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the type of the marker.
    ///
    /// Change the default type of the marker to one of the shapes supported by
    /// Excel.
    ///
    /// # Parameters
    ///
    /// `marker_type`: a {@link ChartMarkerType} enum value.
    #[wasm_bindgen(js_name = "setType", skip_jsdoc)]
    pub fn set_type(&self, marker_type: ChartMarkerType) -> ChartMarker {
        let mut lock = self.inner.lock().unwrap();
        lock.set_type(xlsx::ChartMarkerType::from(marker_type));
        ChartMarker {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the size of the marker.
    ///
    /// Change the default size of the marker.
    ///
    /// # Parameters
    ///
    /// `size` - The size of the marker.
    #[wasm_bindgen(js_name = "setSize", skip_jsdoc)]
    pub fn set_size(&self, size: u8) -> ChartMarker {
        let mut lock = self.inner.lock().unwrap();
        lock.set_size(size);
        ChartMarker {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the formatting properties for a chart marker.
    ///
    /// Set the formatting properties for a chart marker via a {@link ChartFormat}
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
    /// - `set_no_border`: Turn off the border for the chart object.
    ///
    /// # Parameters
    ///
    /// `format`: A {@link ChartFormat} struct reference or a sub struct that will
    /// convert into a `ChartFormat` instance. See the docs for
    /// {@link IntoChartFormat} for details.
    #[wasm_bindgen(js_name = "setFormat", skip_jsdoc)]
    pub fn set_format(&self, format: &ChartFormat) -> ChartMarker {
        let mut lock = self.inner.lock().unwrap();
        lock.set_format(&mut *format.inner.lock().unwrap());
        ChartMarker {
            inner: Arc::clone(&self.inner),
        }
    }
}
