use crate::wrapper::Color;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ShapeFont` struct represents the font format for shape objects.
///
/// Excel uses a standard font dialog for text elements of a shape such as the
/// shape title or axes data labels. It looks like this:
///
/// The {@link ShapeFont} struct represents many of these font options such as font
/// type, size, color and properties such as bold and italic. It is used in
/// conjunction with the {@link Shape#setFont} method.
///
/// It is used in conjunction with the {@link Shape} struct.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ShapeFont {
    pub(crate) inner: Arc<Mutex<xlsx::ShapeFont>>,
}

#[wasm_bindgen]
impl ShapeFont {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ShapeFont {
        ShapeFont {
            inner: Arc::new(Mutex::new(xlsx::ShapeFont::new())),
        }
    }
    #[doc = r" Create a deep clone of this object."]
    #[wasm_bindgen(js_name = "clone")]
    pub fn deep_clone(&self) -> ShapeFont {
        ShapeFont {
            inner: Arc::new(Mutex::new(self.inner.lock().unwrap().clone())),
        }
    }
    /// Set the bold property for the font of a shape element.
    #[wasm_bindgen(js_name = "setBold", skip_jsdoc)]
    pub fn set_bold(&self) -> ShapeFont {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_bold();
        *lock = inner;
        ShapeFont {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the italic property for the font of a shape element.
    #[wasm_bindgen(js_name = "setItalic", skip_jsdoc)]
    pub fn set_italic(&self) -> ShapeFont {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_italic();
        *lock = inner;
        ShapeFont {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the color property for the font of a shape element.
    ///
    /// # Parameters
    ///
    /// - `color`: The font color property defined by a {@link Color} enum
    ///   value.
    #[wasm_bindgen(js_name = "setColor", skip_jsdoc)]
    pub fn set_color(&self, color: Color) -> ShapeFont {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_color(xlsx::Color::from(color));
        *lock = inner;
        ShapeFont {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the shape font name property.
    ///
    /// Set the name/type of a font for a shape element.
    ///
    /// # Parameters
    ///
    /// - `font_name`: The font name property.
    #[wasm_bindgen(js_name = "setName", skip_jsdoc)]
    pub fn set_name(&self, font_name: &str) -> ShapeFont {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_name(font_name);
        *lock = inner;
        ShapeFont {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the size property for the font of a shape element.
    ///
    /// # Parameters
    ///
    /// - `font_size`: The font size property.
    #[wasm_bindgen(js_name = "setSize", skip_jsdoc)]
    pub fn set_size(&self, font_size: f64) -> ShapeFont {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_size(font_size);
        *lock = inner;
        ShapeFont {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the underline property for the font of a shape element.
    ///
    /// The default underline type is the only type supported.
    #[wasm_bindgen(js_name = "setUnderline", skip_jsdoc)]
    pub fn set_underline(&self) -> ShapeFont {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_underline();
        *lock = inner;
        ShapeFont {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the strikethrough property for the font of a shape element.
    #[wasm_bindgen(js_name = "setStrikethrough", skip_jsdoc)]
    pub fn set_strikethrough(&self) -> ShapeFont {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_strikethrough();
        *lock = inner;
        ShapeFont {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Unset the bold property for a font.
    ///
    /// Some shape elements such as titles have a default bold property in
    /// Excel. This method can be used to turn it off.
    #[wasm_bindgen(js_name = "unsetBold", skip_jsdoc)]
    pub fn unset_bold(&self) -> ShapeFont {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.unset_bold();
        *lock = inner;
        ShapeFont {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Display the shape font from right to left for some language support.
    ///
    /// See
    /// {@link Worksheet#setRightToLeft}
    /// for details.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setRightToLeft", skip_jsdoc)]
    pub fn set_right_to_left(&self, enable: bool) -> ShapeFont {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_right_to_left(enable);
        *lock = inner;
        ShapeFont {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the pitch family property for the font of a shape element.
    ///
    /// This function is implemented for completeness but is rarely used in
    /// practice.
    ///
    /// # Parameters
    ///
    /// - `family`: The font family property.
    #[wasm_bindgen(js_name = "setPitchFamily", skip_jsdoc)]
    pub fn set_pitch_family(&self, family: u8) -> ShapeFont {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_pitch_family(family);
        *lock = inner;
        ShapeFont {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the character set property for the font of a shape element.
    ///
    /// Set the font character set. This function is implemented for
    /// completeness but is rarely required in practice.
    ///
    /// # Parameters
    ///
    /// - `character_set`: The font character set property.
    #[wasm_bindgen(js_name = "setCharacterSet", skip_jsdoc)]
    pub fn set_character_set(&self, character_set: u8) -> ShapeFont {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_character_set(character_set);
        *lock = inner;
        ShapeFont {
            inner: Arc::clone(&self.inner),
        }
    }
}
