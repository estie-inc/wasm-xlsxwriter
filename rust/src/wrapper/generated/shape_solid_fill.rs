use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ShapeSolidFill` struct represents a the solid fill for a shape element.
///
/// The {@link ShapeSolidFill} struct represents the formatting properties for the
/// solid fill of a Shape element. In Excel a solid fill is a single color fill
/// without a pattern or gradient.
///
/// `ShapeSolidFill` is a sub property of the {@link ShapeFormat} struct and is used
/// with the {@link ShapeFormat#setSolidFill} method.
///
/// It is used in conjunction with the {@link Shape} struct.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ShapeSolidFill {
    pub(crate) inner: Arc<Mutex<xlsx::ShapeSolidFill>>,
}

#[wasm_bindgen]
impl ShapeSolidFill {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ShapeSolidFill {
        ShapeSolidFill {
            inner: Arc::new(Mutex::new(xlsx::ShapeSolidFill::new())),
        }
    }
}
