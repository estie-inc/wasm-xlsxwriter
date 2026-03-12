use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ChartRange` struct represents a chart range.
///
/// A struct to represent a chart range like `"Sheet1!$A$1:$A$4"`. The struct is
/// public to allow for the {@link IntoChartRange} trait and for a limited set of
/// edge cases, but it isn't generally required to be manipulated by the end
/// user.
///
/// It is used in conjunction with the {@link Chart} struct.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartRange {
    pub(crate) inner: Arc<Mutex<xlsx::ChartRange>>,
}

#[wasm_bindgen]
impl ChartRange {
    #[doc = r" Create a deep clone of this object."]
    #[wasm_bindgen(js_name = "clone")]
    pub fn deep_clone(&self) -> ChartRange {
        ChartRange {
            inner: Arc::new(Mutex::new(self.inner.lock().unwrap().clone())),
        }
    }
}
