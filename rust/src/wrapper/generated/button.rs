use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `Button` struct represents a worksheet button object.
///
/// The `Button` struct is used to create an Excel "Form Control" button object
/// to represent a button on a worksheet.
///
/// The worksheet button object is mainly provided as a way to trigger a VBA
/// macro. See Working with VBA macros for more details. It is
/// used in conjunction with the
/// {@link Worksheet#insertButton} method.
///
/// Note, Button is the only VBA Control supported by `rust_xlsxwriter`. It is
/// unlikely that any other Excel form elements will be added in the future due
/// to the implementation effort required.
///
/// Here is a complete example with a button that has a macro attached to it.
///
/// Output file:
#[derive(Clone)]
#[wasm_bindgen]
pub struct Button {
    pub(crate) inner: Arc<Mutex<xlsx::Button>>,
}

#[wasm_bindgen]
impl Button {
    #[wasm_bindgen(constructor)]
    pub fn new() -> Button {
        Button {
            inner: Arc::new(Mutex::new(xlsx::Button::new())),
        }
    }
    /// Set the button caption.
    ///
    /// The default button caption in Excel is "Button 1", "Button 2" etc. This
    /// method can be used to change that caption to some other text.
    ///
    /// # Parameters
    ///
    /// `caption` - The text to display on the button. It must be less than or
    /// equal to 255 characters.
    #[wasm_bindgen(js_name = "setCaption", skip_jsdoc)]
    pub fn set_caption(&self, caption: &str) -> Button {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_caption(caption);
        *lock = inner;
        Button {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the macro associated with the button.
    ///
    /// The `set_macro()` method can be used to associate an existing VBA macro
    /// with a button object. See Working with VBA macros for
    /// more details on macros in `rust_xlsxwriter`.
    ///
    /// # Parameters
    ///
    /// `name` - The macro name. It should be the same as it appears in the
    /// Excel macros dialog.
    ///
    /// src="https://rustxlsxwriter.github.io/images/button_macro_dialog.png">
    #[wasm_bindgen(js_name = "setMacro", skip_jsdoc)]
    pub fn set_macro(&self, name: &str) -> Button {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_macro(name);
        *lock = inner;
        Button {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the width of the button in pixels.
    ///
    /// # Parameters
    ///
    /// - `width`: The button width in pixels.
    #[wasm_bindgen(js_name = "setWidth", skip_jsdoc)]
    pub fn set_width(&self, width: u32) -> Button {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_width(width);
        *lock = inner;
        Button {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the height of the button in pixels.
    ///
    /// # Parameters
    ///
    /// - `height`: The button height in pixels.
    #[wasm_bindgen(js_name = "setHeight", skip_jsdoc)]
    pub fn set_height(&self, height: u32) -> Button {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_height(height);
        *lock = inner;
        Button {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the alt text for the button to help accessibility.
    ///
    /// The alt text is used with screen readers to help people with visual
    /// disabilities.
    ///
    /// See the following Microsoft documentation on [Everything you need to
    /// know to write effective alt
    /// text](https://support.microsoft.com/en-us/office/everything-you-need-to-know-to-write-effective-alt-text-df98f884-ca3d-456c-807b-1a1fa82f5dc2).
    ///
    /// # Parameters
    ///
    /// - `alt_text`: The alt text string to add to the button.
    #[wasm_bindgen(js_name = "setAltText", skip_jsdoc)]
    pub fn set_alt_text(&self, alt_text: &str) -> Button {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_alt_text(alt_text);
        *lock = inner;
        Button {
            inner: Arc::clone(&self.inner),
        }
    }
}
