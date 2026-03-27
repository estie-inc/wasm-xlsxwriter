use crate::wrapper::Formula;
use crate::wrapper::ObjectMovement;
use crate::wrapper::ShapeFont;
use crate::wrapper::ShapeFormat;
use crate::wrapper::ShapeText;
use crate::wrapper::Url;
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
    /// Set the text in the shape.
    ///
    /// This only applies to shapes that have a textbox option.
    ///
    /// See also {@link Shape#setFont} and {@link Shape#setTextOptions} for
    /// formatting options for text.
    ///
    /// # Parameters
    ///
    /// - `text`: The text for the shape.
    #[wasm_bindgen(js_name = "setText", skip_jsdoc)]
    pub fn set_text(&self, text: &str) -> Shape {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Shape::textbox());
        inner = inner.set_text(text);
        *lock = inner;
        Shape {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the text in the shape from a worksheet cell.
    ///
    /// Set the textbox text from a link to a worksheet cell like `=A1` or
    /// `=Sheet2!A1`.
    ///
    /// This only applies to shapes that have a textbox option.
    ///
    /// # Parameters
    ///
    /// - `cell`: The cell from which the text is linked. Should be a simple
    ///   string or a {@link Formula}.
    #[wasm_bindgen(js_name = "setTextLink", skip_jsdoc)]
    pub fn set_text_link(&self, cell: Formula) -> Shape {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Shape::textbox());
        inner = inner.set_text_link(cell.inner.lock().unwrap().clone());
        *lock = inner;
        Shape {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the width of the shape in pixels.
    ///
    /// The default width for an Excel shape created by `rust_xlsxwriter` is 192
    /// pixels.
    ///
    /// # Parameters
    ///
    /// - `width`: The shape width in pixels. Values less than 5 pixels are
    ///   ignored.
    #[wasm_bindgen(js_name = "setWidth", skip_jsdoc)]
    pub fn set_width(&self, width: u32) -> Shape {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Shape::textbox());
        inner = inner.set_width(width);
        *lock = inner;
        Shape {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the height of the shape in pixels.
    ///
    /// The default height for an Excel shape created by `rust_xlsxwriter` is
    /// 120 pixels.
    ///
    /// See the example above.
    ///
    /// # Parameters
    ///
    /// - `height`: The shape height in pixels. Values less than 5 pixels are
    ///   ignored.
    #[wasm_bindgen(js_name = "setHeight", skip_jsdoc)]
    pub fn set_height(&self, height: u32) -> Shape {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Shape::textbox());
        inner = inner.set_height(height);
        *lock = inner;
        Shape {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the formatting properties for a shape.
    ///
    /// Set the formatting properties for a shape via a {@link ShapeFormat}
    /// object or a sub-struct that implements {@link IntoShapeFormat}.
    ///
    /// The formatting that can be applied via a {@link ShapeFormat} object are:
    ///
    /// - {@link ShapeFormat#setSolidFill}: Set the {@link ShapeSolidFill} properties.
    /// - {@link ShapeFormat#setPatternFill}: Set the {@link ShapePatternFill} properties.
    /// - {@link ShapeFormat#setGradientFill}: Set the {@link ShapeGradientFill} properties.
    /// - {@link ShapeFormat#setNoFill}: Turn off the fill for the shape object.
    /// - {@link ShapeFormat#setLine}: Set the {@link ShapeLine} properties.
    ///   A synonym for {@link ShapeLine} depending on context.
    /// - {@link ShapeFormat#setNoLine}: Turn off the line for the shape object.
    ///
    /// # Parameters
    ///
    /// `format`: A {@link ShapeFormat} struct reference or a sub-struct that will
    /// convert into a `ShapeFormat` instance. See the docs for
    /// {@link IntoShapeFormat} for details.
    #[wasm_bindgen(js_name = "setFormat", skip_jsdoc)]
    pub fn set_format(&self, format: &ShapeFormat) -> Shape {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Shape::textbox());
        inner = inner.set_format(&*format.inner.lock().unwrap());
        *lock = inner;
        Shape {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the font properties of of the shape.
    ///
    /// Set the font properties of a shape text using a {@link ShapeFont} reference.
    /// Example font properties that can be set are:
    ///
    /// - {@link ShapeFont#setBold}
    /// - {@link ShapeFont#setItalic}
    /// - {@link ShapeFont#setColor}
    /// - {@link ShapeFont#setName}
    /// - {@link ShapeFont#setSize}
    /// - {@link ShapeFont#setUnderline}
    /// - {@link ShapeFont#setStrikethrough}
    /// - {@link ShapeFont#setRightToLeft}
    ///
    /// See {@link ShapeFont} for full details.
    ///
    /// # Parameters
    ///
    /// `font`: A {@link ShapeFont} struct reference to represent the font
    /// properties.
    #[wasm_bindgen(js_name = "setFont", skip_jsdoc)]
    pub fn set_font(&self, font: &ShapeFont) -> Shape {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Shape::textbox());
        inner = inner.set_font(&*font.inner.lock().unwrap());
        *lock = inner;
        Shape {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the text option properties of of the shape.
    ///
    /// # Parameters
    ///
    /// - `text_options`: The {@link ShapeText} options.
    #[wasm_bindgen(js_name = "setTextOptions", skip_jsdoc)]
    pub fn set_text_options(&self, text_options: &ShapeText) -> Shape {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Shape::textbox());
        inner = inner.set_text_options(&*text_options.inner.lock().unwrap());
        *lock = inner;
        Shape {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set a Url/Hyperlink for a shape.
    ///
    /// Set a Url/Hyperlink for an shape so that when the user clicks on it they
    /// are redirected to an internal or external location.
    ///
    /// See {@link Url} for an explanation of the URIs supported by Excel and for
    /// other options that can be set.
    ///
    /// # Parameters
    ///
    /// - `link`: The url/hyperlink associate with the shape as a string or
    ///   {@link Url}.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#MaxUrlLengthExceeded} - URL string or anchor exceeds
    ///   Excel's limit of 2080 characters.
    /// - {@link XlsxError#UnknownUrlType} - The URL has an unknown URI type. See
    ///   {@link Worksheet#writeUrl}.
    /// - {@link XlsxError#ParameterError} - URL mouseover tool tip exceeds Excel's
    ///   limit of 255 characters.
    #[wasm_bindgen(js_name = "setUrl", skip_jsdoc)]
    pub fn set_url(&self, link: Url) -> WasmResult<Shape> {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Shape::textbox());
        inner = inner.set_url(link.inner.lock().unwrap().clone())?;
        *lock = inner;
        Ok(Shape {
            inner: Arc::clone(&self.inner),
        })
    }
    /// Set the alt text for the shape to help accessibility.
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
    /// - `alt_text`: The alt text string to add to the shape.
    #[wasm_bindgen(js_name = "setAltText", skip_jsdoc)]
    pub fn set_alt_text(&self, alt_text: &str) -> Shape {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Shape::textbox());
        inner = inner.set_alt_text(alt_text);
        *lock = inner;
        Shape {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the object movement options for a worksheet shape.
    ///
    /// Set the option to define how an shape will behave in Excel if the cells
    /// under the shape are moved, deleted, or have their size changed. In
    /// Excel the options are:
    ///
    /// 1. Move and size with cells.
    /// 2. Move but don't size with cells.
    /// 3. Don't move or size with cells.
    ///
    /// These values are defined in the {@link ObjectMovement} enum.
    ///
    /// The {@link ObjectMovement} enum also provides an additional option to "Move
    /// and size with cells - after the shape is inserted" to allow shapes to
    /// be hidden in rows or columns. In Excel this equates to option 1 above
    /// but the internal shape position calculations are handled differently.
    ///
    /// # Parameters
    ///
    /// - `option`: An shape/object positioning behavior defined by the
    ///   {@link ObjectMovement} enum.
    #[wasm_bindgen(js_name = "setObjectMovement", skip_jsdoc)]
    pub fn set_object_movement(&self, option: ObjectMovement) -> Shape {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Shape::textbox());
        inner = inner.set_object_movement(xlsx::ObjectMovement::from(option));
        *lock = inner;
        Shape {
            inner: Arc::clone(&self.inner),
        }
    }
}
