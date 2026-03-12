use crate::wrapper::ConditionalFormatDateRule;
use crate::wrapper::Format;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ConditionalFormatDate` struct represents a Dates Occurring style
/// conditional format.
///
/// `ConditionalFormatDate` is used to represent a Dates Occurring style
/// conditional format in Excel. This is used to identify dates in ranges like
/// "Last Week" or "Last Month".
///
/// src="https://rustxlsxwriter.github.io/images/conditional_format_date_intro.png">
///
/// For more information see Working with Conditional
/// Formats.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ConditionalFormatDate {
    pub(crate) inner: Arc<Mutex<xlsx::ConditionalFormatDate>>,
}

#[wasm_bindgen]
impl ConditionalFormatDate {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ConditionalFormatDate {
        ConditionalFormatDate {
            inner: Arc::new(Mutex::new(xlsx::ConditionalFormatDate::new())),
        }
    }
    #[doc = r" Create a deep clone of this object."]
    #[wasm_bindgen(js_name = "clone")]
    pub fn deep_clone(&self) -> ConditionalFormatDate {
        ConditionalFormatDate {
            inner: Arc::new(Mutex::new(self.inner.lock().unwrap().clone())),
        }
    }
    /// Set the rule for the Dates Occurring  conditional format rule such
    /// "Today" or "Last Week".
    ///
    /// # Parameters
    ///
    /// - `rule`: A {@link ConditionalFormatDateRule} enum value.
    #[wasm_bindgen(js_name = "setRule", skip_jsdoc)]
    pub fn set_rule(&self, criteria: ConditionalFormatDateRule) -> ConditionalFormatDate {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ConditionalFormatDate::new());
        inner = inner.set_rule(xlsx::ConditionalFormatDateRule::from(criteria));
        *lock = inner;
        ConditionalFormatDate {
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
    pub fn set_format(&self, format: Format) -> ConditionalFormatDate {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ConditionalFormatDate::new());
        inner = inner.set_format(format.inner.lock().unwrap().clone());
        *lock = inner;
        ConditionalFormatDate {
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
    pub fn set_multi_range(&self, range: &str) -> ConditionalFormatDate {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ConditionalFormatDate::new());
        inner = inner.set_multi_range(range);
        *lock = inner;
        ConditionalFormatDate {
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
    pub fn set_stop_if_true(&self, enable: bool) -> ConditionalFormatDate {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ConditionalFormatDate::new());
        inner = inner.set_stop_if_true(enable);
        *lock = inner;
        ConditionalFormatDate {
            inner: Arc::clone(&self.inner),
        }
    }
}
