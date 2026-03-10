use crate::wrapper::ChartGradientStop;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartGradientFill {
    pub(crate) inner: Arc<Mutex<xlsx::ChartGradientFill>>,
}

#[wasm_bindgen]
impl ChartGradientFill {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ChartGradientFill {
        ChartGradientFill {
            inner: Arc::new(Mutex::new(xlsx::ChartGradientFill::new())),
        }
    }
    #[wasm_bindgen(js_name = "setGradientStops", skip_jsdoc)]
    pub fn set_gradient_stops(&self, gradient_stops: Vec<ChartGradientStop>) -> ChartGradientFill {
        let mut lock = self.inner.lock().unwrap();
        lock.set_gradient_stops(
            &gradient_stops
                .iter()
                .map(|x| x.inner.clone())
                .collect::<Vec<_>>(),
        );
        ChartGradientFill {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setAngle", skip_jsdoc)]
    pub fn set_angle(&self, angle: u16) -> ChartGradientFill {
        let mut lock = self.inner.lock().unwrap();
        lock.set_angle(angle);
        ChartGradientFill {
            inner: Arc::clone(&self.inner),
        }
    }
}
