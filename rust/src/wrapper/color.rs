
use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;
use tsify::Tsify;
use serde::{Deserialize, Serialize};

/// The `Color` enum defines Excel colors that can be used throughout the
/// `rust_xlsxwriter` APIs.
/// 
/// There are 3 types of colors within the enum:
/// 
/// 1. Predefined named colors like `Color::Green`.
/// 2. User defined RGB colors such as `Color::RGB(0x4F026A)` using a format
/// similar to html colors like `#RRGGBB`, except as an integer.
/// 3. Theme colors from the standard palette of 60 colors like `Color::Theme(9,
/// 4)`. The theme colors are shown in the image below.
/// 
/// <img
/// src="https://rustxlsxwriter.github.io/images/theme_color_palette.png">
/// 
/// The syntax for theme colors in `Color` is `Theme(color, shade)` where
/// `color` is one of the 0-9 values on the top row and `shade` is the
/// variant in the associated column from 0-5. For example "White, background
/// 1" in the top left is `Theme(0, 0)` and "Orange, Accent 6, Darker 50%" in
/// the bottom right is `Theme(9, 5)`.
/// 
/// Note, there are no plans to support anything other than the default Excel
/// "Office" theme.
#[derive(Debug, Clone, Copy, Default, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum Color {
    RGB(u32),
    Theme(u8, u8),
    #[default]
    Default,
    Automatic,
    Black,
    Blue,
    Brown,
    Cyan,
    Gray,
    Green,
    Lime,
    Magenta,
    Navy,
    Orange,
    Pink,
    Purple,
    Red,
    Silver,
    White,
    Yellow,
}
impl From::<Color> for xlsx::Color {
    fn from(value: Color) -> Self {
        match value {
            Color::RGB(rgb) => xlsx::Color::RGB(rgb),
            Color::Theme(color, shade) => xlsx::Color::Theme(color, shade),
            Color::Default => xlsx::Color::Default,
            Color::Automatic => xlsx::Color::Automatic,
            Color::Black => xlsx::Color::Black,
            Color::Blue => xlsx::Color::Blue,
            Color::Brown => xlsx::Color::Brown,
            Color::Cyan => xlsx::Color::Cyan,
            Color::Gray => xlsx::Color::Gray,
            Color::Green => xlsx::Color::Green,
            Color::Lime => xlsx::Color::Lime,
            Color::Magenta => xlsx::Color::Magenta,
            Color::Navy => xlsx::Color::Navy,
            Color::Orange => xlsx::Color::Orange,
            Color::Pink => xlsx::Color::Pink,
            Color::Purple => xlsx::Color::Purple,
            Color::Red => xlsx::Color::Red,
            Color::Silver => xlsx::Color::Silver,
            Color::White => xlsx::Color::White,
            Color::Yellow => xlsx::Color::Yellow,
        }
    }
}
