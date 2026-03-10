use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ChartPatternFill` struct represents a the pattern fill for a chart
/// element.
///
/// The {@link ChartPatternFill} struct represents the formatting properties for the
/// pattern fill of a Chart element. In Excel a pattern fill is comprised of a
/// simple pixelated pattern and background and foreground colors
///
/// `ChartPatternFill` is a sub property of the {@link ChartFormat} struct and is
/// used with the {@link ChartFormat#setPatternFill} method.
///
/// It is used in conjunction with the {@link Chart} struct.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartPatternFill {
    pub(crate) inner: Arc<Mutex<xlsx::ChartPatternFill>>,
}

#[wasm_bindgen]
impl ChartPatternFill {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ChartPatternFill {
        ChartPatternFill {
            inner: Arc::new(Mutex::new(xlsx::ChartPatternFill::new())),
        }
    }
}
