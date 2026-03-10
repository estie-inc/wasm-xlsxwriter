use crate::wrapper::ShapeGradientFill;
use crate::wrapper::ShapeLine;
use crate::wrapper::ShapePatternFill;
use crate::wrapper::ShapeSolidFill;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ShapeFormat` struct represents formatting for various shape objects.
///
/// Excel uses a standard formatting dialog for the shape elements which
/// generally looks like this:
///
/// The {@link ShapeFormat} struct represents many of these format options and just
/// like Excel it offers a similar formatting interface for a number of the
/// shape sub-elements supported by `rust_xlsxwriter`.
///
/// It is used in conjunction with the {@link Shape} struct.
///
/// The {@link ShapeFormat} struct is generally passed to the `set_format()` method
/// of a shape element. It supports several shape formatting elements such as:
///
/// - {@link ShapeFormat#setSolidFill}: Set the {@link ShapeSolidFill} properties.
/// - {@link ShapeFormat#setPatternFill}: Set the {@link ShapePatternFill}
///   properties.
/// - {@link ShapeFormat#setGradientFill}: Set the {@link ShapeGradientFill}
///   properties.
/// - {@link ShapeFormat#setNoFill}: Turn off the fill for the shape object.
/// - {@link ShapeFormat#setLine}: Set the {@link ShapeLine} properties.
/// - {@link ShapeFormat#setNoLine}: Turn off the line for the shape object.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ShapeFormat {
    pub(crate) inner: Arc<Mutex<xlsx::ShapeFormat>>,
}

#[wasm_bindgen]
impl ShapeFormat {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ShapeFormat {
        ShapeFormat {
            inner: Arc::new(Mutex::new(xlsx::ShapeFormat::new())),
        }
    }
    /// Set the line formatting for a shape element.
    ///
    /// See the {@link ShapeLine} struct for details on the line properties that can
    /// be set.
    ///
    /// # Parameters
    ///
    /// - `line`: A {@link ShapeLine} struct reference.
    #[wasm_bindgen(js_name = "setLine", skip_jsdoc)]
    pub fn set_line(&self, line: ShapeLine) -> ShapeFormat {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_line(&*line.inner.lock().unwrap());
        *lock = inner;
        ShapeFormat {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Turn off the line property for a shape element.
    ///
    /// The line property for a shape element can be turned off if you wish to
    /// hide it.
    #[wasm_bindgen(js_name = "setNoLine", skip_jsdoc)]
    pub fn set_no_line(&self) -> ShapeFormat {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_no_line();
        *lock = inner;
        ShapeFormat {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the solid fill formatting for a shape element.
    ///
    /// See the {@link ShapeSolidFill} struct for details on the solid fill
    /// properties that can be set.
    ///
    /// # Parameters
    ///
    /// - `fill`: A {@link ShapeSolidFill} struct reference.
    #[wasm_bindgen(js_name = "setSolidFill", skip_jsdoc)]
    pub fn set_solid_fill(&self, fill: ShapeSolidFill) -> ShapeFormat {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_solid_fill(&*fill.inner.lock().unwrap());
        *lock = inner;
        ShapeFormat {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Turn off the fill property for a shape element.
    ///
    /// The fill property for a shape element can be turned off if you wish to
    /// hide it and display only the border line.
    #[wasm_bindgen(js_name = "setNoFill", skip_jsdoc)]
    pub fn set_no_fill(&self) -> ShapeFormat {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_no_fill();
        *lock = inner;
        ShapeFormat {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the pattern fill formatting for a shape element.
    ///
    /// See the {@link ShapePatternFill} struct for details on the pattern fill
    /// properties that can be set.
    ///
    /// # Parameters
    ///
    /// - `fill`: A {@link ShapePatternFill} struct reference.
    #[wasm_bindgen(js_name = "setPatternFill", skip_jsdoc)]
    pub fn set_pattern_fill(&self, fill: ShapePatternFill) -> ShapeFormat {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_pattern_fill(&*fill.inner.lock().unwrap());
        *lock = inner;
        ShapeFormat {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the gradient fill formatting for a shape element.
    ///
    /// See the {@link ShapeGradientFill} struct for details on the gradient fill
    /// properties that can be set.
    ///
    /// # Parameters
    ///
    /// - `fill`: A {@link ShapeGradientFill} struct reference.
    #[wasm_bindgen(js_name = "setGradientFill", skip_jsdoc)]
    pub fn set_gradient_fill(&self, fill: ShapeGradientFill) -> ShapeFormat {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_gradient_fill(&*fill.inner.lock().unwrap());
        *lock = inner;
        ShapeFormat {
            inner: Arc::clone(&self.inner),
        }
    }
}
