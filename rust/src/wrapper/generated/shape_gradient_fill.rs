use crate::wrapper::ShapeGradientFillType;
use crate::wrapper::ShapeGradientStop;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct ShapeGradientFill {
    pub(crate) inner: Arc<Mutex<xlsx::ShapeGradientFill>>,
}

#[wasm_bindgen]
impl ShapeGradientFill {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ShapeGradientFill {
        ShapeGradientFill {
            inner: Arc::new(Mutex::new(xlsx::ShapeGradientFill::new())),
        }
    }
    #[wasm_bindgen(js_name = "setType", skip_jsdoc)]
    pub fn set_type(&self, gradient_type: ShapeGradientFillType) -> ShapeGradientFill {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_type(xlsx::ShapeGradientFillType::from(gradient_type));
        *lock = inner;
        ShapeGradientFill {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setGradientStops", skip_jsdoc)]
    pub fn set_gradient_stops(&self, gradient_stops: Vec<ShapeGradientStop>) -> ShapeGradientFill {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_gradient_stops(
            &gradient_stops
                .iter()
                .map(|x| x.inner.lock().unwrap().clone())
                .collect::<Vec<_>>(),
        );
        *lock = inner;
        ShapeGradientFill {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setAngle", skip_jsdoc)]
    pub fn set_angle(&self, angle: u16) -> ShapeGradientFill {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_angle(angle);
        *lock = inner;
        ShapeGradientFill {
            inner: Arc::clone(&self.inner),
        }
    }
}
