use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct Url {
    pub(crate) inner: Arc<Mutex<xlsx::Url>>,
}

#[wasm_bindgen]
impl Url {
    #[wasm_bindgen(constructor)]
    pub fn new(link: &str) -> Url {
        Url {
            inner: Arc::new(Mutex::new(xlsx::Url::new(link))),
        }
    }
}
