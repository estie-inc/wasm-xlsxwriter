use crate::wrapper::ConditionalFormatCustomIcon;
use crate::wrapper::ConditionalFormatIconType;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ConditionalFormatIconSet` struct represents an Icon Set style
/// conditional format.
///
/// `ConditionalFormatIconSet` is used to represent an Icon Set style
/// conditional format in Excel. An Icon Set conditional format highlights items
/// with groups of 3-5 symbols such as traffic lights, arrows, or flags.
///
/// src="https://rustxlsxwriter.github.io/images/conditional_format_icon_intro.png">
///
/// For more information see Working with Conditional
/// Formats.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ConditionalFormatIconSet {
    pub(crate) inner: Arc<Mutex<xlsx::ConditionalFormatIconSet>>,
}

#[wasm_bindgen]
impl ConditionalFormatIconSet {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ConditionalFormatIconSet {
        ConditionalFormatIconSet {
            inner: Arc::new(Mutex::new(xlsx::ConditionalFormatIconSet::new())),
        }
    }
    #[doc = r" Create a deep clone of this object."]
    #[wasm_bindgen(js_name = "clone")]
    pub fn deep_clone(&self) -> ConditionalFormatIconSet {
        ConditionalFormatIconSet {
            inner: Arc::new(Mutex::new(self.inner.lock().unwrap().clone())),
        }
    }
    /// Set the icon types such as traffic lights or histograms.
    ///
    /// # Parameters
    ///
    /// - `icon_type`: A {@link ConditionalFormatIconType} enum value.
    #[wasm_bindgen(js_name = "setIconType", skip_jsdoc)]
    pub fn set_icon_type(&self, icon_type: ConditionalFormatIconType) -> ConditionalFormatIconSet {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ConditionalFormatIconSet::new());
        inner = inner.set_icon_type(xlsx::ConditionalFormatIconType::from(icon_type));
        *lock = inner;
        ConditionalFormatIconSet {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Reverse the order of icons from lowest to highest.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "reverseIcons", skip_jsdoc)]
    pub fn reverse_icons(&self, enable: bool) -> ConditionalFormatIconSet {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ConditionalFormatIconSet::new());
        inner = inner.reverse_icons(enable);
        *lock = inner;
        ConditionalFormatIconSet {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Show only the icons and not the data in the cells.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "showIconsOnly", skip_jsdoc)]
    pub fn show_icons_only(&self, enable: bool) -> ConditionalFormatIconSet {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ConditionalFormatIconSet::new());
        inner = inner.show_icons_only(enable);
        *lock = inner;
        ConditionalFormatIconSet {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set user defined rules for the icon set.
    ///
    /// Set custom rules for the icon ranges in an Icon Set conditional format.
    ///
    /// Excel sets default rules for Icon Set conditional formats depending on
    /// whether they have 3, 4, or 5 icons. The equivalent rules in `rust_xlsxwriter` are:
    ///
    /// # Parameters
    ///
    /// `icons`: A slice of {@link ConditionalFormatCustomIcon} objects.
    #[wasm_bindgen(js_name = "setIcons", skip_jsdoc)]
    pub fn set_icons(&self, icons: Vec<ConditionalFormatCustomIcon>) -> ConditionalFormatIconSet {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ConditionalFormatIconSet::new());
        inner = inner.set_icons(
            &icons
                .iter()
                .map(|x| x.inner.lock().unwrap().clone())
                .collect::<Vec<_>>(),
        );
        *lock = inner;
        ConditionalFormatIconSet {
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
    pub fn set_multi_range(&self, range: &str) -> ConditionalFormatIconSet {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ConditionalFormatIconSet::new());
        inner = inner.set_multi_range(range);
        *lock = inner;
        ConditionalFormatIconSet {
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
    pub fn set_stop_if_true(&self, enable: bool) -> ConditionalFormatIconSet {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ConditionalFormatIconSet::new());
        inner = inner.set_stop_if_true(enable);
        *lock = inner;
        ConditionalFormatIconSet {
            inner: Arc::clone(&self.inner),
        }
    }
}
