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
}
