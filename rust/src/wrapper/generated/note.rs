use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `Note` struct represents a worksheet note object.
///
/// A Note is a post-it style message that is revealed when the user hovers over
/// a worksheet cell. The presence of a Note is indicated by a small red
/// triangle in the upper right-hand corner of the cell.
///
/// Notes are used in conjunction with the
/// {@link Worksheet#insertNote} method.
///
/// In versions of Excel prior to Office 365, Notes were referred to as
/// "Comments". The name "Comment" is now used for a newer style threaded comment,
/// and "Note" is used for the older non-threaded version. See the Microsoft docs
/// on [The difference between threaded comments and notes].
///
/// [The difference between threaded comments and notes]:
///     https://support.microsoft.com/en-us/office/the-difference-between-threaded-comments-and-notes-75a51eec-4092-42ab-abf8-7669077b7be3
///
/// Note, the newer Threaded Comments are unlikely to be added to
/// `rust_xlsxwriter` due to the fact that the feature relies on company
/// specific metadata to identify the comment author.
#[derive(Clone)]
#[wasm_bindgen]
pub struct Note {
    pub(crate) inner: Arc<Mutex<xlsx::Note>>,
}

#[wasm_bindgen]
impl Note {
    #[wasm_bindgen(constructor)]
    pub fn new(text: &str) -> Note {
        Note {
            inner: Arc::new(Mutex::new(xlsx::Note::new(text))),
        }
    }
    /// Reset the text in the note.
    ///
    /// In general, the text of the note is set in the {@link Note#new}
    /// constructor, but if required, you can use the `reset_text()` method to
    /// reset the text for a note. This allows a single `Note` instance to be
    /// used multiple times and avoids the small overhead of creating a new
    /// instance each time.
    ///
    /// # Parameters
    ///
    /// - `text`: The text for the note.
    #[wasm_bindgen(js_name = "resetText", skip_jsdoc)]
    pub fn reset_text(&self, text: &str) -> Note {
        let mut lock = self.inner.lock().unwrap();
        lock.reset_text(text);
        Note {
            inner: Arc::clone(&self.inner),
        }
    }
}
