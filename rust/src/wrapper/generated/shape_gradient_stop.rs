use crate::wrapper::Color;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ShapeGradientStop` struct represents a gradient fill data point.
///
/// The {@link ShapeGradientStop} struct represents the properties of a data point
/// (a stop) that is used to generate a gradient fill.
///
/// src="https://rustxlsxwriter.github.io/images/gradient_fill_options.png">
///
/// Excel supports between 2 and 10 gradient stops which define the a color and
/// its position in the gradient as a percentage. These colors and positions
/// are used to interpolate a gradient fill.
///
/// Gradient formats are generally used with the
/// {@link ShapeGradientFill#setGradientStops} method and
/// {@link ShapeGradientFill}.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ShapeGradientStop {
    pub(crate) inner: Arc<Mutex<xlsx::ShapeGradientStop>>,
}

#[wasm_bindgen]
impl ShapeGradientStop {
    #[wasm_bindgen(constructor)]
    pub fn new(color: Color, position: u8) -> ShapeGradientStop {
        ShapeGradientStop {
            inner: Arc::new(Mutex::new(xlsx::ShapeGradientStop::new(
                xlsx::Color::from(color),
                position,
            ))),
        }
    }
    #[doc = r" Create a deep clone of this object."]
    #[wasm_bindgen(js_name = "clone")]
    pub fn deep_clone(&self) -> ShapeGradientStop {
        ShapeGradientStop {
            inner: Arc::new(Mutex::new(self.inner.lock().unwrap().clone())),
        }
    }
}
