use crate::wrapper::Color;
use crate::wrapper::ShapeLineDashType;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ShapeLine` struct represents a shape line/border.
///
/// The {@link ShapeLine} struct represents the formatting properties for the line
/// of a Shape element. It is a sub property of the {@link ShapeFormat} struct and
/// is used with the {@link ShapeFormat#setLine} method.
///
/// For 2D shapes the line property usually represents the border.
///
/// It is used in conjunction with the {@link Shape} struct.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ShapeLine {
    pub(crate) inner: Arc<Mutex<xlsx::ShapeLine>>,
}

#[wasm_bindgen]
impl ShapeLine {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ShapeLine {
        ShapeLine {
            inner: Arc::new(Mutex::new(xlsx::ShapeLine::new())),
        }
    }
    #[doc = r" Create a deep clone of this object."]
    #[wasm_bindgen(js_name = "clone")]
    pub fn deep_clone(&self) -> ShapeLine {
        ShapeLine {
            inner: Arc::new(Mutex::new(self.inner.lock().unwrap().clone())),
        }
    }
    /// Set the color of a line/border.
    ///
    /// # Parameters
    ///
    /// - `color`: The color property defined by a {@link Color} enum value or a
    ///   type that can convert `Into` a {@link Color} such as html like string.
    #[wasm_bindgen(js_name = "setColor", skip_jsdoc)]
    pub fn set_color(&self, color: Color) -> ShapeLine {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ShapeLine::new());
        inner = inner.set_color(xlsx::Color::from(color));
        *lock = inner;
        ShapeLine {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the width of the line or border.
    ///
    /// # Parameters
    ///
    /// - `width`: The width should be specified in increments of 0.25 of a
    ///   point as in Excel. The width can be an number type that convert
    ///   `Into` `f64`. The default width is 0.75.
    #[wasm_bindgen(js_name = "setWidth", skip_jsdoc)]
    pub fn set_width(&self, width: f64) -> ShapeLine {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ShapeLine::new());
        inner = inner.set_width(width);
        *lock = inner;
        ShapeLine {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the dash type of the line or border.
    ///
    /// # Parameters
    ///
    /// - `dash_type`: A {@link ShapeLineDashType} enum value.
    #[wasm_bindgen(js_name = "setDashType", skip_jsdoc)]
    pub fn set_dash_type(&self, dash_type: ShapeLineDashType) -> ShapeLine {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ShapeLine::new());
        inner = inner.set_dash_type(xlsx::ShapeLineDashType::from(dash_type));
        *lock = inner;
        ShapeLine {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the transparency of a line/border.
    ///
    /// Set the transparency of a line/border for a Shape element. You must also
    /// specify a line color in order for the transparency to be applied.
    ///
    /// # Parameters
    ///
    /// - `transparency`: The color transparency in the range 0 <= transparency
    ///   <= 100. The default value is 0.
    #[wasm_bindgen(js_name = "setTransparency", skip_jsdoc)]
    pub fn set_transparency(&self, transparency: u8) -> ShapeLine {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ShapeLine::new());
        inner = inner.set_transparency(transparency);
        *lock = inner;
        ShapeLine {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the shape line as hidden.
    ///
    /// The method is sometimes required to turn off a default line type in
    /// order to highlight some other part of the line. This can also be
    /// achieved more succinctly using the {@link ShapeFormat#setNoLine}
    /// method.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off (not hidden) by default.
    #[wasm_bindgen(js_name = "setHidden", skip_jsdoc)]
    pub fn set_hidden(&self, enable: bool) -> ShapeLine {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ShapeLine::new());
        inner = inner.set_hidden(enable);
        *lock = inner;
        ShapeLine {
            inner: Arc::clone(&self.inner),
        }
    }
}
