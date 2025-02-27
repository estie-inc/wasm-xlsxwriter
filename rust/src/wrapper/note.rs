use std::sync::{Arc, Mutex};

use rust_xlsxwriter::{self as xlsx};
use wasm_bindgen::prelude::*;
#[derive(Clone)]
#[wasm_bindgen]
pub struct Note {
    pub(crate) inner: Arc<Mutex<xlsx::Note>>,
}

#[wasm_bindgen]
impl Note {
    pub(crate) fn lock(&self) -> std::sync::MutexGuard<'_, xlsx::Note> {
        self.inner.lock().unwrap()
    }

    /// Create a new Note object to represent an Excel cell note.
    /// The text of the Note is added in the constructor.
    /// @param {string} text - The text that will appear in the note.
    /// @returns {Note} - The note object.
    #[wasm_bindgen(constructor, skip_jsdoc)]
    pub fn new(text: &str) -> Note {
        Note {
            inner: Arc::new(Mutex::new(xlsx::Note::new(text))),
        }
    }
}
