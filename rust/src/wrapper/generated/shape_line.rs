use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ShapeLine` struct represents a shape line/border.
///
/// The {@link ShapeLine} struct represents the formatting properties for the line
/// of a Shape element. It is a sub property of the {@link ShapeFormat} struct and
/// is used with the {@link ShapeFormat#setLine} method.
///
/// For 2D shapes the line property usually represents the border.
///
/// It is used in conjunction with the {@link Shape} struct.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ShapeLine {
    pub(crate) inner: Arc<Mutex<xlsx::ShapeLine>>,
}

#[wasm_bindgen]
impl ShapeLine {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ShapeLine {
        ShapeLine {
            inner: Arc::new(Mutex::new(xlsx::ShapeLine::new())),
        }
    }
}
