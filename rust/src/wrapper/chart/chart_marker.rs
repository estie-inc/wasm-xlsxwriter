use rust_xlsxwriter::{self as xlsx};
use wasm_bindgen::prelude::*;

use crate::wrapper::chart::chart_format::ChartFormat;
use crate::wrapper::chart::chart_marker_type::ChartMarkerType;

#[wasm_bindgen]
pub struct ChartMarker {
    pub(crate) inner: xlsx::ChartMarker,
}

#[wasm_bindgen]
impl ChartMarker {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ChartMarker {
        ChartMarker {
            inner: xlsx::ChartMarker::new(),
        }
    }

    /// Set the automatic/default marker type.
    ///
    /// Allow the marker type to be set automatically by Excel.
    ///
    /// @return {ChartMarker} - The ChartMarker instance.
    #[wasm_bindgen(js_name = "setAutomatic")]
    pub fn set_automatic(mut self) -> ChartMarker {
        self.inner.set_automatic();
        self
    }

    /// Set the formatting properties for a chart marker.
    ///
    /// @param {ChartFormat} format - The chart format properties.
    /// @return {ChartMarker} - The ChartMarker instance.
    #[wasm_bindgen(js_name = "setFormat")]
    pub fn set_format(mut self, format: &mut ChartFormat) -> ChartMarker {
        self.inner.set_format(&mut format.inner);
        self
    }

    /// Turn off/hide a chart marker.
    ///
    /// This method can be used to turn off markers for an individual data series
    /// in a chart that has default markers for all series.
    ///
    /// @return {ChartMarker} - The ChartMarker instance.
    #[wasm_bindgen(js_name = "setNone")]
    pub fn set_none(mut self) -> ChartMarker {
        self.inner.set_none();
        self
    }

    /// Set the marker size.
    ///
    /// @param {number} size - The marker size.
    /// @return {ChartMarker} - The ChartMarker instance.
    #[wasm_bindgen(js_name = "setSize")]
    pub fn set_size(mut self, size: u8) -> ChartMarker {
        self.inner.set_size(size);
        self
    }

    /// Set the marker type.
    ///
    /// @param {number} marker_type - The marker type.
    /// @return {ChartMarker} - The ChartMarker instance.
    #[wasm_bindgen(js_name = "setType")]
    pub fn set_type(mut self, marker_type: ChartMarkerType) -> ChartMarker {
        self.inner.set_type(marker_type.into());
        self
    }
}
