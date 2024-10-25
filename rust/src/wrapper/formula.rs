use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct Formula {
    pub(crate) inner: xlsx::Formula,
}

#[wasm_bindgen]
impl Formula {
    #[wasm_bindgen(constructor)]
    pub fn new(formula: &str) -> Formula {
        Formula {
            inner: xlsx::Formula::new(formula),
        }
    }

    #[wasm_bindgen(js_name = "setResult")]
    pub fn set_result(&self, result: &str) -> Formula {
        Formula {
            inner: self.clone().inner.set_result(result),
        }
    }
}
