use crate::wrapper::Color;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ChartGradientStop` struct represents a gradient fill data point.
///
/// The {@link ChartGradientStop} struct represents the properties of a data point
/// (a stop) that is used to generate a gradient fill.
///
/// src="https://rustxlsxwriter.github.io/images/gradient_fill_options.png">
///
/// Excel supports between 2 and 10 gradient stops which define the a color and
/// its position in the gradient as a percentage. These colors and positions
/// are used to interpolate a gradient fill.
///
/// Gradient formats are generally used with the
/// {@link ChartGradientFill#setGradientStops} method and
/// {@link ChartGradientFill}.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartGradientStop {
    pub(crate) inner: Arc<Mutex<xlsx::ChartGradientStop>>,
}

#[wasm_bindgen]
impl ChartGradientStop {
    #[wasm_bindgen(constructor)]
    pub fn new(color: Color, position: u8) -> ChartGradientStop {
        ChartGradientStop {
            inner: Arc::new(Mutex::new(xlsx::ChartGradientStop::new(
                color.inner,
                position,
            ))),
        }
    }
}
