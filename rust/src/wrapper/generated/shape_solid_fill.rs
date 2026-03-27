use crate::wrapper::Color;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ShapeSolidFill` struct represents a the solid fill for a shape element.
///
/// The {@link ShapeSolidFill} struct represents the formatting properties for the
/// solid fill of a Shape element. In Excel a solid fill is a single color fill
/// without a pattern or gradient.
///
/// `ShapeSolidFill` is a sub property of the {@link ShapeFormat} struct and is used
/// with the {@link ShapeFormat#setSolidFill} method.
///
/// It is used in conjunction with the {@link Shape} struct.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ShapeSolidFill {
    pub(crate) inner: Arc<Mutex<xlsx::ShapeSolidFill>>,
}

#[wasm_bindgen]
impl ShapeSolidFill {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ShapeSolidFill {
        ShapeSolidFill {
            inner: Arc::new(Mutex::new(xlsx::ShapeSolidFill::new())),
        }
    }
    #[doc = r" Create a deep clone of this object."]
    #[wasm_bindgen(js_name = "clone")]
    pub fn deep_clone(&self) -> ShapeSolidFill {
        ShapeSolidFill {
            inner: Arc::new(Mutex::new(self.inner.lock().unwrap().clone())),
        }
    }
    /// Set the color of a solid fill.
    ///
    /// # Parameters
    ///
    /// - `color`: The color property defined by a {@link Color} enum value or
    ///   a type that can convert `Into` a {@link Color}.
    #[wasm_bindgen(js_name = "setColor", skip_jsdoc)]
    pub fn set_color(&self, color: Color) -> ShapeSolidFill {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ShapeSolidFill::new());
        inner = inner.set_color(xlsx::Color::from(color));
        *lock = inner;
        ShapeSolidFill {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the transparency of a solid fill.
    ///
    /// Set the transparency of a solid fill color for a Shape element. You must
    /// also specify a line color in order for the transparency to be applied.
    ///
    /// # Parameters
    ///
    /// - `transparency`: The color transparency in the range 0 <= transparency
    ///   <= 100. The default value is 0.
    #[wasm_bindgen(js_name = "setTransparency", skip_jsdoc)]
    pub fn set_transparency(&self, transparency: u8) -> ShapeSolidFill {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ShapeSolidFill::new());
        inner = inner.set_transparency(transparency);
        *lock = inner;
        ShapeSolidFill {
            inner: Arc::clone(&self.inner),
        }
    }
}
