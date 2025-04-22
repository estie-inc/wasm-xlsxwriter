use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[wasm_bindgen]
pub enum HeaderImagePosition {
    Left,
    Center,
    Right,
}

impl From<HeaderImagePosition> for xlsx::HeaderImagePosition {
    fn from(position: HeaderImagePosition) -> Self {
        match position {
            HeaderImagePosition::Left => xlsx::HeaderImagePosition::Left,
            HeaderImagePosition::Center => xlsx::HeaderImagePosition::Center,
            HeaderImagePosition::Right => xlsx::HeaderImagePosition::Right,
        }
    }
}
