use crate::wrapper::Color;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ChartFont` struct represents the font format for various chart objects.
///
/// Excel uses a standard font dialog for text elements of a chart such as the
/// chart title or axes data labels. It looks like this:
///
/// The {@link ChartFont} struct represents many of these font options such as font
/// type, size, color and properties such as bold and italic. It is generally
/// used in conjunction with a `set_font()` method for a chart element.
///
/// It is used in conjunction with the {@link Chart} struct.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartFont {
    pub(crate) inner: Arc<Mutex<xlsx::ChartFont>>,
}

#[wasm_bindgen]
impl ChartFont {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ChartFont {
        ChartFont {
            inner: Arc::new(Mutex::new(xlsx::ChartFont::new())),
        }
    }
    #[doc = r" Create a deep clone of this object."]
    #[wasm_bindgen(js_name = "clone")]
    pub fn deep_clone(&self) -> ChartFont {
        ChartFont {
            inner: Arc::new(Mutex::new(self.inner.lock().unwrap().clone())),
        }
    }
    /// Set the bold property for the font of a chart element.
    #[wasm_bindgen(js_name = "setBold", skip_jsdoc)]
    pub fn set_bold(&self) -> ChartFont {
        let mut lock = self.inner.lock().unwrap();
        lock.set_bold();
        ChartFont {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the italic property for the font of a chart element.
    #[wasm_bindgen(js_name = "setItalic", skip_jsdoc)]
    pub fn set_italic(&self) -> ChartFont {
        let mut lock = self.inner.lock().unwrap();
        lock.set_italic();
        ChartFont {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the color property for the font of a chart element.
    ///
    /// # Parameters
    ///
    /// - `color`: The font color property defined by a {@link Color} enum
    ///   value.
    #[wasm_bindgen(js_name = "setColor", skip_jsdoc)]
    pub fn set_color(&self, color: Color) -> ChartFont {
        let mut lock = self.inner.lock().unwrap();
        lock.set_color(xlsx::Color::from(color));
        ChartFont {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the chart font name property.
    ///
    /// Set the name/type of a font for a chart element.
    ///
    /// # Parameters
    ///
    /// - `font_name`: The font name property.
    #[wasm_bindgen(js_name = "setName", skip_jsdoc)]
    pub fn set_name(&self, font_name: &str) -> ChartFont {
        let mut lock = self.inner.lock().unwrap();
        lock.set_name(font_name);
        ChartFont {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the size property for the font of a chart element.
    ///
    /// # Parameters
    ///
    /// - `font_size`: The font size property.
    #[wasm_bindgen(js_name = "setSize", skip_jsdoc)]
    pub fn set_size(&self, font_size: f64) -> ChartFont {
        let mut lock = self.inner.lock().unwrap();
        lock.set_size(font_size);
        ChartFont {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the text rotation property for the font of a chart element.
    ///
    /// Set the rotation angle of the text in a cell. The rotation can be any
    /// angle in the range -90 to 90 degrees, or 270-271 to indicate text where
    /// the letters run from top to bottom, see below.
    ///
    /// # Parameters
    ///
    /// - `rotation`: The rotation angle in the range `-90 <= rotation <= 90`.
    ///   Two special case values are supported:
    ///   - `270`: Stacked text, where the text runs from top to bottom.
    ///   - `271`: A special variant of stacked text for East Asian fonts.
    #[wasm_bindgen(js_name = "setRotation", skip_jsdoc)]
    pub fn set_rotation(&self, rotation: i16) -> ChartFont {
        let mut lock = self.inner.lock().unwrap();
        lock.set_rotation(rotation);
        ChartFont {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the underline property for the font of a chart element.
    ///
    /// The default underline type is the only type supported.
    #[wasm_bindgen(js_name = "setUnderline", skip_jsdoc)]
    pub fn set_underline(&self) -> ChartFont {
        let mut lock = self.inner.lock().unwrap();
        lock.set_underline();
        ChartFont {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the strikethrough property for the font of a chart element.
    #[wasm_bindgen(js_name = "setStrikethrough", skip_jsdoc)]
    pub fn set_strikethrough(&self) -> ChartFont {
        let mut lock = self.inner.lock().unwrap();
        lock.set_strikethrough();
        ChartFont {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Unset the bold property for a font.
    ///
    /// Some chart elements such as titles have a default bold property in
    /// Excel. This method can be used to turn it off.
    #[wasm_bindgen(js_name = "unsetBold", skip_jsdoc)]
    pub fn unset_bold(&self) -> ChartFont {
        let mut lock = self.inner.lock().unwrap();
        lock.unset_bold();
        ChartFont {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Display the chart font from right to left for some language support.
    ///
    /// See
    /// {@link Worksheet#setRightToLeft}
    /// for details.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setRightToLeft", skip_jsdoc)]
    pub fn set_right_to_left(&self, enable: bool) -> ChartFont {
        let mut lock = self.inner.lock().unwrap();
        lock.set_right_to_left(enable);
        ChartFont {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the pitch family property for the font of a chart element.
    ///
    /// This function is implemented for completeness but is rarely used in
    /// practice.
    ///
    /// # Parameters
    ///
    /// - `family`: The font family property.
    #[wasm_bindgen(js_name = "setPitchFamily", skip_jsdoc)]
    pub fn set_pitch_family(&self, family: u8) -> ChartFont {
        let mut lock = self.inner.lock().unwrap();
        lock.set_pitch_family(family);
        ChartFont {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the character set property for the font of a chart element.
    ///
    /// Set the font character set. This function is implemented for
    /// completeness but is rarely required in practice.
    ///
    /// # Parameters
    ///
    /// - `character_set`: The font character set property.
    #[wasm_bindgen(js_name = "setCharacterSet", skip_jsdoc)]
    pub fn set_character_set(&self, character_set: u8) -> ChartFont {
        let mut lock = self.inner.lock().unwrap();
        lock.set_character_set(character_set);
        ChartFont {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the default bold property for the font.
    ///
    /// The is mainly only required for testing to ensure strict compliance with
    /// Excel's output.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setDefaultBold", skip_jsdoc)]
    pub fn set_default_bold(&self, enable: bool) -> ChartFont {
        let mut lock = self.inner.lock().unwrap();
        lock.set_default_bold(enable);
        ChartFont {
            inner: Arc::clone(&self.inner),
        }
    }
}
