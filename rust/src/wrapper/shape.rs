use std::sync::{Arc, Mutex};

use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

use crate::wrapper::Shape;

#[wasm_bindgen]
impl Shape {
    #[wasm_bindgen(constructor)]
    pub fn textbox() -> Shape {
        Shape {
            inner: Arc::new(Mutex::new(xlsx::Shape::textbox())),
        }
    }
}
