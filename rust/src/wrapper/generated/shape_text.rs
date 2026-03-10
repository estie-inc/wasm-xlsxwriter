use crate::wrapper::ShapeTextDirection;
use crate::wrapper::ShapeTextHorizontalAlignment;
use crate::wrapper::ShapeTextVerticalAlignment;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ShapeText` struct represents the text options for a shape element.
///
/// The {@link ShapeText} struct represents the text option properties for a Shape
///  element:
///
/// src="https://rustxlsxwriter.github.io/images/shape_text_options_dialog.png">
///
/// Currently only the vertical, horizontal and text direction properties are
/// supported.
///
/// `ShapeText` is a sub property of the {@link ShapeFormat} struct and is used with
/// the {@link Shape#setTextOptions} method. See also {@link ShapeFont}.
///
/// It is used in conjunction with the {@link Shape} struct.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ShapeText {
    pub(crate) inner: Arc<Mutex<xlsx::ShapeText>>,
}

#[wasm_bindgen]
impl ShapeText {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ShapeText {
        ShapeText {
            inner: Arc::new(Mutex::new(xlsx::ShapeText::new())),
        }
    }
    /// Set the horizontal alignment for the text in a shape textbox.
    ///
    /// This method sets the horizontal alignment for the text in a shape while
    /// {@link ShapeText#setVerticalAlignment} sets the alignment for the text
    /// bounding box. See the example below.
    ///
    /// # Parameters
    ///
    /// - `alignment`: A {@link ShapeTextHorizontalAlignment} enum value.
    #[wasm_bindgen(js_name = "setHorizontalAlignment", skip_jsdoc)]
    pub fn set_horizontal_alignment(&self, alignment: ShapeTextHorizontalAlignment) -> ShapeText {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_horizontal_alignment(xlsx::ShapeTextHorizontalAlignment::from(alignment));
        *lock = inner;
        ShapeText {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the vertical alignment for the textbox in a shape.
    ///
    /// This method sets the vertical alignment of the textbox in a shape while
    /// {@link ShapeText#setHorizontalAlignment} sets the alignment for the
    /// text within the textbox. See the example above.
    ///
    /// # Parameters
    ///
    /// - `alignment`: A {@link ShapeTextVerticalAlignment} enum value.
    #[wasm_bindgen(js_name = "setVerticalAlignment", skip_jsdoc)]
    pub fn set_vertical_alignment(&self, alignment: ShapeTextVerticalAlignment) -> ShapeText {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_vertical_alignment(xlsx::ShapeTextVerticalAlignment::from(alignment));
        *lock = inner;
        ShapeText {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the text direction of the text in the text box.
    ///
    /// This is useful for languages that display text vertically.
    ///
    /// # Parameters
    ///
    /// - `direction`: The {@link ShapeTextDirection} of the text.
    #[wasm_bindgen(js_name = "setDirection", skip_jsdoc)]
    pub fn set_direction(&self, direction: ShapeTextDirection) -> ShapeText {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_direction(xlsx::ShapeTextDirection::from(direction));
        *lock = inner;
        ShapeText {
            inner: Arc::clone(&self.inner),
        }
    }
}
