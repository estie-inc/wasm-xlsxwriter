use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

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
#[derive(Debug, Copy, Clone, Default)]
#[wasm_bindgen]
pub struct Color {
    pub(crate) inner: xlsx::Color,
}

impl Color {
    pub fn new(inner: xlsx::Color) -> Self {
        Color { inner }
    }
}

#[wasm_bindgen]
impl Color {
    /// The default color for an Excel property.
    #[wasm_bindgen(constructor)]
    pub fn default() -> Color {
        Color::new(xlsx::Color::Default)
    }

    /// A user defined RGB color in the range 0x000000 (black) to 0xFFFFFF
    /// (white). Any values outside this range will be ignored with a a warning.
    #[wasm_bindgen]
    pub fn rgb(hex: u32) -> Color {
        Color::new(xlsx::Color::RGB(hex))
    }

    /// A theme color on the default palette (see the image above). The syntax
    /// for theme colors is `Theme(color, shade)` where `color` is one of the
    /// 0-9 values on the top row and `shade` is the variant in the associated
    /// column from 0-5. Any values outside these ranges will be ignored with a
    /// a warning.
    #[wasm_bindgen]
    pub fn theme(color: u8, shade: u8) -> Color {
        Color::new(xlsx::Color::Theme(color, shade))
    }

    /// The Automatic color for an Excel property. This is usually the same as
    /// the `Default` color but can vary according to system settings.
    #[wasm_bindgen]
    pub fn automatic() -> Color {
        Color::new(xlsx::Color::Automatic)
    }

    /// Convert from a Html style color string line "#6495ED" into a {@link Color} enum value.
    #[wasm_bindgen]
    pub fn parse(s: &str) -> Color {
        let color = xlsx::Color::from(s);
        Color { inner: color }
    }

    /// The color Black with a RGB value of 0x000000.
    #[wasm_bindgen]
    pub fn black() -> Color {
        Color::new(xlsx::Color::Black)
    }

    /// The color Blue with a RGB value of 0x0000FF.
    #[wasm_bindgen]
    pub fn blue() -> Color {
        Color::new(xlsx::Color::Blue)
    }

    /// The color Brown with a RGB value of 0x800000.
    #[wasm_bindgen]
    pub fn brown() -> Color {
        Color::new(xlsx::Color::Brown)
    }

    /// The color Cyan with a RGB value of 0x00FFFF.
    #[wasm_bindgen]
    pub fn cyan() -> Color {
        Color::new(xlsx::Color::Cyan)
    }

    /// The color Gray with a RGB value of 0x808080.
    #[wasm_bindgen]
    pub fn gray() -> Color {
        Color::new(xlsx::Color::Gray)
    }

    /// The color Green with a RGB value of 0x008000.
    #[wasm_bindgen]
    pub fn green() -> Color {
        Color::new(xlsx::Color::Green)
    }

    /// The color Lime with a RGB value of 0x00FF00.
    #[wasm_bindgen]
    pub fn lime() -> Color {
        Color::new(xlsx::Color::Lime)
    }

    /// The color Magenta with a RGB value of 0xFF00FF.
    #[wasm_bindgen]
    pub fn magenta() -> Color {
        Color::new(xlsx::Color::Magenta)
    }

    /// The color Navy with a RGB value of 0x000080.
    #[wasm_bindgen]
    pub fn navy() -> Color {
        Color::new(xlsx::Color::Navy)
    }

    /// The color Orange with a RGB value of 0xFF6600.
    #[wasm_bindgen]
    pub fn orange() -> Color {
        Color::new(xlsx::Color::Orange)
    }

    /// The color Pink with a RGB value of 0xFFC0CB.
    #[wasm_bindgen]
    pub fn pink() -> Color {
        Color::new(xlsx::Color::Pink)
    }

    /// The color Purple with a RGB value of 0x800080.
    #[wasm_bindgen]
    pub fn purple() -> Color {
        Color::new(xlsx::Color::Purple)
    }

    /// The color Red with a RGB value of 0xFF0000.
    #[wasm_bindgen]
    pub fn red() -> Color {
        Color::new(xlsx::Color::Red)
    }

    /// The color Silver with a RGB value of 0xC0C0C0.
    #[wasm_bindgen]
    pub fn silver() -> Color {
        Color::new(xlsx::Color::Silver)
    }

    /// The color White with a RGB value of 0xFFFFFF.
    #[wasm_bindgen]
    pub fn white() -> Color {
        Color::new(xlsx::Color::White)
    }

    /// The color Yellow with a RGB value of 0xFFFF00
    #[wasm_bindgen]
    pub fn yellow() -> Color {
        Color::new(xlsx::Color::Yellow)
    }
}
