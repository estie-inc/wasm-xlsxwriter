use rust_xlsxwriter::{self as xlsx};
use wasm_bindgen::prelude::*;

use super::{color::Color, format::Format};
use crate::macros::wrap_struct;

#[derive(Clone, Copy, Debug)]
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

wrap_struct!(
    Note,
    xlsx::Note,
    set_author(name: &str),
    add_author_prefix(enable: bool),
    reset_text(text: &str),
    set_width(width: u32),
    set_height(height: u32),
    set_visible(enable: bool),
    set_background_color(color: Color),
    set_font_name(font_name: &str),
    set_font_size(font_size: f64),
    set_font_family(font_family: u8),
    set_format(format: Format),
    set_alt_text(alt_text: &str),
    set_object_movement(option: ObjectMovement)
);
