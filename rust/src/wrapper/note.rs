use std::sync::Arc;

use rust_xlsxwriter::{self as xlsx};
use wasm_bindgen::prelude::*;

use super::{color::Color, format::Format};

use crate::wrap_struct;

wrap_struct!(
    Note,
    xlsx::Note,
    new(text: &str),
    set_author(name: &str),
    add_author_prefix(enable: bool),
    set_width(width: u32),
    set_height(height: u32),
    set_visible(enable: bool),
    set_font_name(font_name: &str),
    set_font_size(font_size: f64),
    set_font_family(font_family: u8),
    set_alt_text(alt_text: &str)
);

// 特殊なケースを個別に実装
#[wasm_bindgen]
impl Note {
    #[wasm_bindgen(js_name = "resetText", skip_jsdoc)]
    pub fn reset_text(&self, text: &str) -> Note {
        let mut note = self.lock();
        let _ = note.reset_text(text);
        Note {
            inner: Arc::clone(&self.inner),
        }
    }

    #[wasm_bindgen(js_name = "setBackgroundColor", skip_jsdoc)]
    pub fn set_background_color(&self, color: Color) -> Note {
        let mut lock = self.lock();
        let mut inner = lock.clone();
        inner = inner.set_background_color(color.inner);
        *lock = inner;
        Note {
            inner: Arc::clone(&self.inner),
        }
    }

    #[doc(hidden)]
    #[wasm_bindgen(js_name = "setFormat", skip_jsdoc)]
    pub fn set_format(&self, format: Format) -> Note {
        let mut lock = self.lock();
        let mut inner = lock.clone();
        inner = inner.set_format(&*format.lock());
        *lock = inner;
        Note {
            inner: Arc::clone(&self.inner),
        }
    }

    #[wasm_bindgen(js_name = "setObjectMovement", skip_jsdoc)]
    pub fn set_object_movement(&self, option: ObjectMovement) -> Note {
        let mut lock = self.lock();
        let mut inner = lock.clone();
        inner = inner.set_object_movement(option.into());
        *lock = inner;
        Note {
            inner: Arc::clone(&self.inner),
        }
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
