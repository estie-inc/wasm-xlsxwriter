use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `Color` enum defines Excel colors that can be used throughout the
/// `rust_xlsxwriter` APIs.
///
/// There are 3 types of colors within the enum:
///
/// 1. Predefined named colors like `Color::Green`.
/// 2. User defined RGB colors such as `Color::RGB(0x4F026A)` using a format
///    similar to html colors like `#RRGGBB`, except as an integer.
/// 3. Theme colors from the standard palette of 60 colors like `Color::Theme(9,
///    4)`. The theme colors are shown in the image below.
///
///    <img
///    src="https://rustxlsxwriter.github.io/images/theme_color_palette.png">
///
///    The syntax for theme colors in `Color` is `Theme(color, shade)` where
///    `color` is one of the 0-9 values on the top row and `shade` is the
///    variant in the associated column from 0-5. For example "White, background
///    1" in the top left is `Theme(0, 0)` and "Orange, Accent 6, Darker 50%" in
///    the bottom right is `Theme(9, 5)`.
///
/// Note, there are no plans to support anything other than the default Excel
/// "Office" theme.
#[derive(Debug, Clone, Serialize, Deserialize, Tsify, Default)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum Color {
    /// A user-defined RGB color in the range 0x000000 (black) to 0xFFFFFF
    /// (white). Any values outside this range will be ignored with a warning.
    RGB(u32),
    /// A theme color on the default palette (see the image above). The syntax
    /// for theme colors is `Theme(color, shade)` where `color` is one of the
    /// 0-9 values on the top row and `shade` is the variant in the associated
    /// column from 0-5. Any values outside these ranges will be ignored with a
    /// warning.
    Theme(u8, u8),
    /// The default color for an Excel property.
    #[default]
    Default,
    /// The Automatic color for an Excel property. This is usually the same as
    /// the `Default` color but can vary according to system settings.
    Automatic,
    /// The color Black with an RGB value of 0x000000.
    Black,
    /// The color Blue with an RGB value of 0x0000FF.
    Blue,
    /// The color Brown with an RGB value of 0x800000.
    Brown,
    /// The color Cyan with an RGB value of 0x00FFFF.
    Cyan,
    /// The color Gray with an RGB value of 0x808080.
    Gray,
    /// The color Green with an RGB value of 0x008000.
    Green,
    /// The color Lime with an RGB value of 0x00FF00.
    Lime,
    /// The color Magenta with an RGB value of 0xFF00FF.
    Magenta,
    /// The color Navy with an RGB value of 0x000080.
    Navy,
    /// The color Orange with an RGB value of 0xFF6600.
    Orange,
    /// The color Pink with an RGB value of 0xFFC0CB.
    Pink,
    /// The color Purple with an RGB value of 0x800080.
    Purple,
    /// The color Red with an RGB value of 0xFF0000.
    Red,
    /// The color Silver with an RGB value of 0xC0C0C0.
    Silver,
    /// The color White with an RGB value of 0xFFFFFF.
    White,
    /// The color Yellow with an RGB value of 0xFFFF00.
    Yellow,
}

impl From<Color> for xlsx::Color {
    fn from(value: Color) -> xlsx::Color {
        match value {
            Color::RGB(v0) => xlsx::Color::RGB(v0),
            Color::Theme(v0, v1) => xlsx::Color::Theme(v0, v1),
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
