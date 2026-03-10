use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ObjectMovement` enum defines the movement of worksheet objects such as
/// images and charts.
///
/// This enum defines the way control a worksheet object such as [Image],
/// Chart(crate::Chart), Note(crate::Note), Shape(crate::Shape) or
/// Button(crate::Button) moves when the cells underneath it are moved,
/// resized or deleted. This equates to the following Excel options:
///
/// <img src="https://rustxlsxwriter.github.io/images/object_movement.png">
///
/// Used with Image.setObjectMovement().
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
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
    fn from(value: ObjectMovement) -> xlsx::ObjectMovement {
        match value {
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
