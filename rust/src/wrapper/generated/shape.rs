use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `Shape` struct represents a worksheet shape object.
///
/// Currently, the only Excel shape type that is implemented is the `Textbox`
/// shape:
///
/// Output file:
///
/// See also the {@link Worksheet#insertShape}
/// and
/// {@link Worksheet#insertShapeWithOffset}
/// methods. Note that it isn't possible to insert textboxes into other
/// `rust_xlsxwriter` objects such as {@link Chart}.
///
/// ## Support for other Excel shape types
///
/// Currently, the only Excel shape type that is supported is the `Textbox`
/// shape.
///
/// The internal structure of {@link Shape} and the associated XML-generating code
/// is structured to support other shape types, but none are currently
/// implemented. The rationale for this is:
///
/// - Unlike applications like `PowerPoint`, the shape object is not widely used
///   in Excel.
/// - The most common shape used in Excel is the Textbox/Rectangle.
/// - Alternative ways of displaying information, such as {@link Image}
///   or {@link Note}, are already supported.
/// - Each shape or connector type requires a significant number of test cases
///   to verify their functionality and interaction.
///
/// The last is the main reason that I don't wish to support other shape types.
/// The implementation burden is small, but the test and maintenance burden is
/// large. As such, I won't accept Pull Requests to add more shape types.
/// However, I will leave the door open for feature requests that provide a
/// justification.
#[derive(Clone)]
#[wasm_bindgen]
pub struct Shape {
    pub(crate) inner: Arc<Mutex<xlsx::Shape>>,
}

#[wasm_bindgen]
impl Shape {
    #[doc = r" Create a deep clone of this object."]
    #[wasm_bindgen(js_name = "clone")]
    pub fn deep_clone(&self) -> Shape {
        Shape {
            inner: Arc::new(Mutex::new(self.inner.lock().unwrap().clone())),
        }
    }
}
