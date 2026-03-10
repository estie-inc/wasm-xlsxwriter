use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ConditionalFormatCustomIcon` struct represents an icon in an Icon Set
/// style conditional format.
///
/// The `ConditionalFormatCustomIcon` struct is create user defined icons for a
/// {@link ConditionalFormatIconSet} conditional format.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ConditionalFormatCustomIcon {
    pub(crate) inner: Arc<Mutex<xlsx::ConditionalFormatCustomIcon>>,
}

#[wasm_bindgen]
impl ConditionalFormatCustomIcon {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ConditionalFormatCustomIcon {
        ConditionalFormatCustomIcon {
            inner: Arc::new(Mutex::new(xlsx::ConditionalFormatCustomIcon::new())),
        }
    }
}
