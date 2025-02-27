use std::sync::{Arc, Mutex};

use rust_xlsxwriter::{self as xlsx};
use wasm_bindgen::prelude::*;

use super::{color::Color, format::Format};
#[derive(Clone)]
#[wasm_bindgen]
pub struct Note {
    pub(crate) inner: Arc<Mutex<xlsx::Note>>,
}

macro_rules! impl_method {
    ($self:ident.$method:ident($($arg:expr),*)) => {
        let mut lock = $self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Note::new(""));
        inner = inner.$method($($arg),*);
        let _ = std::mem::replace(&mut *lock, inner);
        return Note {
            inner: Arc::clone(&$self.inner),
        }
    };
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
    #[wasm_bindgen(js_name = "setAuthor", skip_jsdoc)]
    pub fn set_author(&self, name: &str) -> Note {
        impl_method!(self.set_author(name));
    }
    #[wasm_bindgen(js_name = "addAuthorPrefix", skip_jsdoc)]
    pub fn add_author_prefix(&self, enable: bool) -> Note {
        impl_method!(self.add_author_prefix(enable));
    }
    #[wasm_bindgen(js_name = "resetText", skip_jsdoc)]
    pub fn reset_text(&self, text: &str) -> Note {
        let mut note = self.inner.lock().unwrap();
        let _ = note.reset_text(text);
        Note {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setWidth", skip_jsdoc)]
    pub fn set_width(&self, width: u32) -> Note {
        impl_method!(self.set_width(width));
    }
    #[wasm_bindgen(js_name = "setHeight", skip_jsdoc)]
    pub fn set_height(&self, height: u32) -> Note {
        impl_method!(self.set_height(height));
    }
    #[wasm_bindgen(js_name = "setVisible", skip_jsdoc)]
    pub fn set_visible(&self, enable: bool) -> Note {
        impl_method!(self.set_visible(enable));
    }
    #[wasm_bindgen(js_name = "setBackgroundColor", skip_jsdoc)]
    pub fn set_background_color(&self, color: Color) -> Note {
        impl_method!(self.set_background_color(color.inner));
    }
    #[wasm_bindgen(js_name = "setFontName", skip_jsdoc)]
    pub fn set_font_name(&self, font_name: &str) -> Note {
        impl_method!(self.set_font_name(font_name));
    }
    #[wasm_bindgen(js_name = "setFontSize", skip_jsdoc)]
    pub fn set_font_size(&self, font_size: f64) -> Note {
        impl_method!(self.set_font_size(font_size));
    }
    #[doc(hidden)]
    #[wasm_bindgen(js_name = "setFontFamily", skip_jsdoc)]
    pub fn set_font_family(&self, font_family: u8) -> Note {
        impl_method!(self.set_font_family(font_family));
    }
    #[doc(hidden)]
    #[wasm_bindgen(js_name = "setFormat", skip_jsdoc)]
    pub fn set_format(&self, format: Format) -> Note {
        impl_method!(self.set_format(&*format.lock()));
    }
    #[wasm_bindgen(js_name = "setAltText", skip_jsdoc)]
    pub fn set_alt_text(&self, alt_text: &str) -> Note {
        impl_method!(self.set_alt_text(alt_text));
    }
    #[wasm_bindgen(js_name = "setObjectMovement", skip_jsdoc)]
    pub fn set_object_movement(&self, option: ObjectMovement) -> Note {
        impl_method!(self.set_object_movement(option.into()));
    }
}

#[derive(Clone, Debug, PartialEq, Eq, Copy)]
#[wasm_bindgen]
pub enum ObjectMovement {
    /// Move and size the worksheet object with the cells. Default for charts.
    MoveAndSizeWithCells,

    /// Move but don't size the worksheet object with the cells. Default for
    /// images.
    MoveButDontSizeWithCells,

    /// Don't move or size the worksheet object with the cells.
    DontMoveOrSizeWithCells,

    /// Same as `MoveAndSizeWithCells` except hidden cells are applied after the
    /// object is inserted. This allows the insertion of objects into hidden
    /// rows or columns.
    MoveAndSizeWithCellsAfter,
}

impl From<ObjectMovement> for xlsx::ObjectMovement {
    fn from(underline: ObjectMovement) -> xlsx::ObjectMovement {
        match underline {
            ObjectMovement::MoveAndSizeWithCells => xlsx::ObjectMovement::MoveAndSizeWithCells,
            ObjectMovement::MoveButDontSizeWithCells => {
                xlsx::ObjectMovement::MoveButDontSizeWithCells
            }
            ObjectMovement::DontMoveOrSizeWithCells => {
                xlsx::ObjectMovement::DontMoveOrSizeWithCells
            }
            ObjectMovement::MoveAndSizeWithCellsAfter => {
                xlsx::ObjectMovement::MoveAndSizeWithCellsAfter
            }
        }
    }
}
