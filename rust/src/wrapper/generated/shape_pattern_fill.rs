use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ShapePatternFill` struct represents a the pattern fill for a shape
/// element.
///
/// The {@link ShapePatternFill} struct represents the formatting properties for the
/// pattern fill of a Shape element. In Excel a pattern fill is comprised of a
/// simple pixelated pattern and background and foreground colors
///
/// `ShapePatternFill` is a sub property of the {@link ShapeFormat} struct and is
/// used with the {@link ShapeFormat#setPatternFill} method.
///
/// It is used in conjunction with the {@link Shape} struct.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ShapePatternFill {
    pub(crate) inner: Arc<Mutex<xlsx::ShapePatternFill>>,
}

#[wasm_bindgen]
impl ShapePatternFill {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ShapePatternFill {
        ShapePatternFill {
            inner: Arc::new(Mutex::new(xlsx::ShapePatternFill::new())),
        }
    }
}
