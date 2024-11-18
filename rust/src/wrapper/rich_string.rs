use std::sync::{Arc, Mutex, MutexGuard};

use super::format::Format;
use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

/// A rich string is a string that can have multiple formats.
#[derive(Clone)]
#[wasm_bindgen]
pub struct RichString {
    inner: Arc<Mutex<Vec<(xlsx::Format, String)>>>,
}

#[wasm_bindgen]
impl RichString {
    pub(crate) fn lock(&self) -> MutexGuard<Vec<(xlsx::Format, String)>> {
        self.inner.lock().unwrap()
    }

    /// Create a new RichString struct.
    ///
    /// @returns {RichString} - The rich string object.
    #[wasm_bindgen(constructor, skip_jsdoc)]
    pub fn new() -> RichString {
        RichString {
            inner: Arc::new(Mutex::new(vec![])),
        }
    }

    /// Add a new part to the rich string.
    #[wasm_bindgen]
    pub fn append(&self, format: &Format, string: String) -> RichString {
        let mut inner = self.inner.lock().unwrap();
        inner.push((format.lock().clone(), string));
        RichString {
            inner: Arc::clone(&self.inner),
        }
    }
}
