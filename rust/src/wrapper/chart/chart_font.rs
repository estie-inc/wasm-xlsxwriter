use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

use crate::wrapper::color::Color;

/// The `ChartFont` struct represents a chart font.
///
/// The `ChartFont` struct is used to define the font properties for chart elements
/// such as chart titles, axis labels, data labels and other text elements in a chart.
///
/// It is used in conjunction with the {@link Chart} struct.
#[wasm_bindgen]
#[derive(Clone)]
pub struct ChartFont {
    pub(crate) inner: xlsx::ChartFont,
}

#[wasm_bindgen]
impl ChartFont {
    /// Create a new `ChartFont` object.
    #[wasm_bindgen(constructor)]
    pub fn new() -> ChartFont {
        ChartFont {
            inner: xlsx::ChartFont::new(),
        }
    }

    /// Set the font bold property.
    ///
    /// @return {ChartFont} - The ChartFont instance.
    #[wasm_bindgen(js_name = "setBold")]
    pub fn set_bold(&mut self) -> ChartFont {
        self.inner.set_bold();
        self.clone()
    }

    /// Set the font character set value.
    ///
    /// Set the font character set value using the standard Windows values.
    /// This is generally only required when using non-standard fonts.
    ///
    /// @param {number} charset - The font character set value.
    /// @return {ChartFont} - The ChartFont instance.
    #[wasm_bindgen(js_name = "setCharacterSet")]
    pub fn set_character_set(&mut self, character_set: u8) -> ChartFont {
        self.inner.set_character_set(character_set);
        self.clone()
    }

    /// Set the font color property.
    ///
    /// @param {Color} color - The font color property.
    /// @return {ChartFont} - The ChartFont instance.
    #[wasm_bindgen(js_name = "setColor")]
    pub fn set_color(&mut self, color: &Color) -> ChartFont {
        self.inner.set_color(color.inner);
        self.clone()
    }

    /// Set the font italic property.
    ///
    /// @return {ChartFont} - The ChartFont instance.
    #[wasm_bindgen(js_name = "setItalic")]
    pub fn set_italic(&mut self) -> ChartFont {
        self.inner.set_italic();
        self.clone()
    }

    /// Set the font name/family property.
    ///
    /// Set the font name for the chart element. Excel can only display fonts that
    /// are installed on the system that it is running on. Therefore it is generally
    /// best to use standard Excel fonts.
    ///
    /// @param {string} name - The font name property.
    /// @return {ChartFont} - The ChartFont instance.
    #[wasm_bindgen(js_name = "setName")]
    pub fn set_name(&mut self, name: &str) -> ChartFont {
        self.inner.set_name(name);
        self.clone()
    }

    /// Set the font pitch and family value.
    ///
    /// Set the font pitch and family value using the standard Windows values.
    /// This is generally only required when using non-standard fonts.
    ///
    /// @param {number} pitch_family - The font pitch and family value.
    /// @return {ChartFont} - The ChartFont instance.
    #[wasm_bindgen(js_name = "setPitchFamily")]
    pub fn set_pitch_family(&mut self, pitch_family: u8) -> ChartFont {
        self.inner.set_pitch_family(pitch_family);
        self.clone()
    }

    /// Set the right to left property.
    ///
    /// Set the right to left property. This is generally only required when using
    /// non-standard fonts.
    ///
    /// @param {boolean} enable - Turn the property on/off. Defaults to true.
    /// @return {ChartFont} - The ChartFont instance.
    #[wasm_bindgen(js_name = "setRightToLeft")]
    pub fn set_right_to_left(&mut self, enable: bool) -> ChartFont {
        self.inner.set_right_to_left(enable);
        self.clone()
    }

    /// Set the font rotation angle.
    ///
    /// Set the rotation angle of the font text in the range -90 to 90 degrees.
    ///
    /// @param {number} rotation - The rotation angle in degrees.
    /// @return {ChartFont} - The ChartFont instance.
    #[wasm_bindgen(js_name = "setRotation")]
    pub fn set_rotation(&mut self, rotation: i16) -> ChartFont {
        self.inner.set_rotation(rotation);
        self.clone()
    }

    /// Set the font size property.
    ///
    /// Set the font size for the chart element.
    ///
    /// @param {number} size - The font size property.
    /// @return {ChartFont} - The ChartFont instance.
    #[wasm_bindgen(js_name = "setSize")]
    pub fn set_size(&mut self, size: f64) -> ChartFont {
        self.inner.set_size(size);
        self.clone()
    }

    /// Set the font strikethrough property.
    ///
    /// Set the strikethrough property. This is generally only required when using
    /// non-standard fonts.
    ///
    /// @return {ChartFont} - The ChartFont instance.
    #[wasm_bindgen(js_name = "setStrikethrough")]
    pub fn set_strikethrough(&mut self) -> ChartFont {
        self.inner.set_strikethrough();
        self.clone()
    }

    /// Set the font underline property.
    ///
    /// Set the font underline. This is generally only required when using
    /// non-standard fonts.
    ///
    /// @param {number} underline - The font underline value.
    /// @return {ChartFont} - The ChartFont instance.
    #[wasm_bindgen(js_name = "setUnderline")]
    pub fn set_underline(&mut self) -> ChartFont {
        self.inner.set_underline();
        self.clone()
    }

    /// Unset bold property.
    ///
    /// Unset the bold property. This is generally only required when using
    /// non-standard fonts.
    ///
    /// @return {ChartFont} - The ChartFont instance.
    #[wasm_bindgen(js_name = "unsetBold")]
    pub fn unset_bold(&mut self) -> ChartFont {
        self.inner.unset_bold();
        self.clone()
    }
}
