use crate::wrapper::Color;
use crate::wrapper::ConditionalFormatType;
use crate::wrapper::ConditionalFormatValue;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ConditionalFormat2ColorScale` struct represents a 2 Color Scale
/// conditional format.
///
/// `ConditionalFormat2ColorScale` is used to represent a Cell style conditional
/// format in Excel. A 2 Color Scale Cell conditional format shows a per cell
/// color gradient from the minimum value to the maximum value.
///
/// src="https://rustxlsxwriter.github.io/images/conditional_format_2color_intro.png">
///
/// For more information see Working with Conditional
/// Formats.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ConditionalFormat2ColorScale {
    pub(crate) inner: Arc<Mutex<xlsx::ConditionalFormat2ColorScale>>,
}

#[wasm_bindgen]
impl ConditionalFormat2ColorScale {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ConditionalFormat2ColorScale {
        ConditionalFormat2ColorScale {
            inner: Arc::new(Mutex::new(xlsx::ConditionalFormat2ColorScale::new())),
        }
    }
    #[doc = r" Create a deep clone of this object."]
    #[wasm_bindgen(js_name = "clone")]
    pub fn deep_clone(&self) -> ConditionalFormat2ColorScale {
        ConditionalFormat2ColorScale {
            inner: Arc::new(Mutex::new(self.inner.lock().unwrap().clone())),
        }
    }
    /// Set the type and value of the minimum in the 2 color scale.
    ///
    /// Set the minimum type (number, percent, formula or percentile) and value
    /// for a 2 color scale type of conditional format. By default the minimum
    /// is the lowest value in the conditional formatting range.
    ///
    /// # Parameters
    ///
    /// - `rule_type`: A {@link ConditionalFormatType} enum value.
    /// - `value`: Any type that can convert into a {@link ConditionalFormatValue}
    ///   such as numbers, dates, times and formula ranges. String values are
    ///   ignored in this type of conditional format.
    #[wasm_bindgen(js_name = "setMinimum", skip_jsdoc)]
    pub fn set_minimum(
        &self,
        rule_type: ConditionalFormatType,
        value: ConditionalFormatValue,
    ) -> ConditionalFormat2ColorScale {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ConditionalFormat2ColorScale::new());
        inner = inner.set_minimum(
            xlsx::ConditionalFormatType::from(rule_type),
            value.inner.lock().unwrap().clone(),
        );
        *lock = inner;
        ConditionalFormat2ColorScale {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the type and value of the maximum in the 2 color scale.
    ///
    /// Set the maximum type (number, percent, formula or percentile) and value
    /// for a 2 color scale type of conditional format. By default the maximum
    /// is the highest value in the conditional formatting range.
    ///
    /// # Parameters
    ///
    /// - `rule_type`: A {@link ConditionalFormatType} enum value.
    /// - `value`: Any type that can convert into a {@link ConditionalFormatValue}
    ///   such as numbers, dates, times and formula ranges. String values are
    ///   ignored in this type of conditional format.
    #[wasm_bindgen(js_name = "setMaximum", skip_jsdoc)]
    pub fn set_maximum(
        &self,
        rule_type: ConditionalFormatType,
        value: ConditionalFormatValue,
    ) -> ConditionalFormat2ColorScale {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ConditionalFormat2ColorScale::new());
        inner = inner.set_maximum(
            xlsx::ConditionalFormatType::from(rule_type),
            value.inner.lock().unwrap().clone(),
        );
        *lock = inner;
        ConditionalFormat2ColorScale {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the color of the minimum in the 2 color scale.
    ///
    /// Set the minimum color value for a 2 color scale type of conditional
    /// format. By default the minimum color is `#FFEF9C` (yellow).
    ///
    /// # Parameters
    ///
    /// - `color`: The color property defined by a {@link Color} enum value or a
    ///   type that can convert `Into` a {@link Color}.
    #[wasm_bindgen(js_name = "setMinimumColor", skip_jsdoc)]
    pub fn set_minimum_color(&self, color: Color) -> ConditionalFormat2ColorScale {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ConditionalFormat2ColorScale::new());
        inner = inner.set_minimum_color(xlsx::Color::from(color));
        *lock = inner;
        ConditionalFormat2ColorScale {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the color of the maximum in the 2 color scale.
    ///
    /// Set the maximum color value for a 2 color scale type of conditional
    /// format. By default the maximum color is `#63BE7B` (green).
    ///
    /// # Parameters
    ///
    /// - `color`: The color property defined by a {@link Color} enum value or a
    ///   type that can convert `Into` a {@link Color}.
    #[wasm_bindgen(js_name = "setMaximumColor", skip_jsdoc)]
    pub fn set_maximum_color(&self, color: Color) -> ConditionalFormat2ColorScale {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ConditionalFormat2ColorScale::new());
        inner = inner.set_maximum_color(xlsx::Color::from(color));
        *lock = inner;
        ConditionalFormat2ColorScale {
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
    pub fn set_multi_range(&self, range: &str) -> ConditionalFormat2ColorScale {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ConditionalFormat2ColorScale::new());
        inner = inner.set_multi_range(range);
        *lock = inner;
        ConditionalFormat2ColorScale {
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
    pub fn set_stop_if_true(&self, enable: bool) -> ConditionalFormat2ColorScale {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ConditionalFormat2ColorScale::new());
        inner = inner.set_stop_if_true(enable);
        *lock = inner;
        ConditionalFormat2ColorScale {
            inner: Arc::clone(&self.inner),
        }
    }
}
