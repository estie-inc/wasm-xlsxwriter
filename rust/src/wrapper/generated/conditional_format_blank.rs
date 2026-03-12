use crate::wrapper::Format;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ConditionalFormatBlank` struct represents a Blank/Non-blank conditional
/// format.
///
/// `ConditionalFormatBlank` is used to represent a Blank or Non-blank style
/// conditional format in Excel. A Blank conditional format highlights blank
/// values in a range while the inverted version highlights non-blanks values.
///
/// src="https://rustxlsxwriter.github.io/images/conditional_format_blank_intro.png">
///
/// For more information see Working with Conditional
/// Formats.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ConditionalFormatBlank {
    pub(crate) inner: Arc<Mutex<xlsx::ConditionalFormatBlank>>,
}

#[wasm_bindgen]
impl ConditionalFormatBlank {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ConditionalFormatBlank {
        ConditionalFormatBlank {
            inner: Arc::new(Mutex::new(xlsx::ConditionalFormatBlank::new())),
        }
    }
    #[doc = r" Create a deep clone of this object."]
    #[wasm_bindgen(js_name = "clone")]
    pub fn deep_clone(&self) -> ConditionalFormatBlank {
        ConditionalFormatBlank {
            inner: Arc::new(Mutex::new(self.inner.lock().unwrap().clone())),
        }
    }
    /// Invert the functionality of the conditional format to get Non-blank
    /// values instead of Blank values.
    ///
    /// See the example above.
    #[wasm_bindgen(js_name = "invert", skip_jsdoc)]
    pub fn invert(&self) -> ConditionalFormatBlank {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ConditionalFormatBlank::new());
        inner = inner.invert();
        *lock = inner;
        ConditionalFormatBlank {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the {@link Format} of the conditional format rule.
    ///
    /// Set the {@link Format} that will be applied to the cell range if the conditional
    /// format rule applies. Not all cell format properties can be set in a
    /// conditional format. See Excel's limitations on conditional format
    /// properties for
    /// more information.
    ///
    /// See the examples above.
    ///
    /// # Parameters
    ///
    /// - `format`: The {@link Format} property for the conditional format.
    #[wasm_bindgen(js_name = "setFormat", skip_jsdoc)]
    pub fn set_format(&self, format: Format) -> ConditionalFormatBlank {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ConditionalFormatBlank::new());
        inner = inner.set_format(format.inner.lock().unwrap().clone());
        *lock = inner;
        ConditionalFormatBlank {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set an additional multi-cell range for the conditional format.
    ///
    /// The `set_multi_range()` method is used to extend a conditional
    /// format over non-contiguous ranges like `"B3:D6 I3:K6 B9:D12
    /// I9:K12"`.
    ///
    /// See Selecting a non-contiguous
    /// range
    /// for more information.
    ///
    /// # Parameters
    ///
    /// - `range`: A string like type representing an Excel range.
    ///
    ///   Note, you can use an Excel range like `"$B$3:$D$6,$I$3:$K$6"` or
    ///   omit the `$` anchors and replace the commas with spaces to have a
    ///   clearer range like `"B3:D6 I3:K6"`. The documentation and examples
    ///   use the latter format for clarity but it you are copying and
    ///   pasting from Excel you can use the first format.
    ///
    ///   Note, if the range is invalid then Excel will omit it silently.
    #[wasm_bindgen(js_name = "setMultiRange", skip_jsdoc)]
    pub fn set_multi_range(&self, range: &str) -> ConditionalFormatBlank {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ConditionalFormatBlank::new());
        inner = inner.set_multi_range(range);
        *lock = inner;
        ConditionalFormatBlank {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the "Stop if True" option for the conditional format rule.
    ///
    /// The `set_stop_if_true()` method can be used to set the “Stop if true”
    /// feature of a conditional formatting rule when more than one rule is
    /// applied to a cell or a range of cells. When this parameter is set then
    /// subsequent rules are not evaluated if the current rule is true.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setStopIfTrue", skip_jsdoc)]
    pub fn set_stop_if_true(&self, enable: bool) -> ConditionalFormatBlank {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ConditionalFormatBlank::new());
        inner = inner.set_stop_if_true(enable);
        *lock = inner;
        ConditionalFormatBlank {
            inner: Arc::clone(&self.inner),
        }
    }
}
