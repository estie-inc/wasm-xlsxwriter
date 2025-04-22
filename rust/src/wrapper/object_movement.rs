use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

/// The `ObjectMovement` enum defines how a worksheet object (like an image or chart) will behave
/// if the cells under the object are moved, deleted, or have their size changed.
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
    fn from(movement: ObjectMovement) -> xlsx::ObjectMovement {
        match movement {
            ObjectMovement::MoveAndSizeWithCells => xlsx::ObjectMovement::MoveAndSizeWithCells,
            ObjectMovement::MoveButDontSizeWithCells => xlsx::ObjectMovement::MoveButDontSizeWithCells,
            ObjectMovement::DontMoveOrSizeWithCells => xlsx::ObjectMovement::DontMoveOrSizeWithCells,
            ObjectMovement::MoveAndSizeWithCellsAfter => xlsx::ObjectMovement::MoveAndSizeWithCellsAfter,
        }
    }
} 