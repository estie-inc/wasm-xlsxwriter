use crate::wrapper::Color;
use crate::wrapper::ConditionalFormatDataBarAxisPosition;
use crate::wrapper::ConditionalFormatDataBarDirection;
use crate::wrapper::ConditionalFormatType;
use crate::wrapper::ConditionalFormatValue;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ConditionalFormatDataBar` struct represents a Data Bar conditional
/// format.
///
/// `ConditionalFormatDataBar` is used to represent a Cell style conditional
/// format in Excel. A Data Bar Cell conditional format shows a per cell color
/// gradient from the minimum value to the maximum value.
///
/// src="https://rustxlsxwriter.github.io/images/conditional_format_databar_intro.png">
///
/// The options that can be applied to a data bar conditional format in Excel
/// are shown in the image below.
///
/// src="https://rustxlsxwriter.github.io/images/conditional_format_databar_intro2.png">
///
/// The methods to replicate these options are explained in the following
/// documentation.
///
/// For more information see Working with Conditional
/// Formats.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ConditionalFormatDataBar {
    pub(crate) inner: Arc<Mutex<xlsx::ConditionalFormatDataBar>>,
}

#[wasm_bindgen]
impl ConditionalFormatDataBar {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ConditionalFormatDataBar {
        ConditionalFormatDataBar {
            inner: Arc::new(Mutex::new(xlsx::ConditionalFormatDataBar::new())),
        }
    }
    #[doc = r" Create a deep clone of this object."]
    #[wasm_bindgen(js_name = "clone")]
    pub fn deep_clone(&self) -> ConditionalFormatDataBar {
        ConditionalFormatDataBar {
            inner: Arc::new(Mutex::new(self.inner.lock().unwrap().clone())),
        }
    }
    /// Set the type and value of the minimum in the data bar.
    ///
    /// Set the minimum type (number, percent, formula or percentile) and value
    /// for a data bar conditional format. By default the minimum is set
    /// automatically.
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
    ) -> ConditionalFormatDataBar {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ConditionalFormatDataBar::new());
        inner = inner.set_minimum(
            xlsx::ConditionalFormatType::from(rule_type),
            value.inner.lock().unwrap().clone(),
        );
        *lock = inner;
        ConditionalFormatDataBar {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the type and value of the maximum in the data bar.
    ///
    /// Set the maximum type (number, percent, formula or percentile) and value
    /// for a data bar conditional format. By default the maximum is set
    /// automatically.
    ///
    /// See the example above.
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
    ) -> ConditionalFormatDataBar {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ConditionalFormatDataBar::new());
        inner = inner.set_maximum(
            xlsx::ConditionalFormatType::from(rule_type),
            value.inner.lock().unwrap().clone(),
        );
        *lock = inner;
        ConditionalFormatDataBar {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the color of the fill in the data bar.
    ///
    /// Set the fill color for a data bar conditional format. By default the
    /// data bar color is `#638EC6` (blue).
    ///
    /// # Parameters
    ///
    /// - `color`: The color property defined by a {@link Color} enum value or a
    ///   type that can convert `Into` a {@link Color}.
    #[wasm_bindgen(js_name = "setFillColor", skip_jsdoc)]
    pub fn set_fill_color(&self, color: Color) -> ConditionalFormatDataBar {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ConditionalFormatDataBar::new());
        inner = inner.set_fill_color(xlsx::Color::from(color));
        *lock = inner;
        ConditionalFormatDataBar {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the color of the border in the data bar.
    ///
    /// Set the border color for a data bar conditional format. By default the
    /// border is the same color as the data bar: `#638EC6` (blue).
    ///
    /// # Parameters
    ///
    /// - `color`: The color property defined by a {@link Color} enum value or a
    ///   type that can convert `Into` a {@link Color}.
    #[wasm_bindgen(js_name = "setBorderColor", skip_jsdoc)]
    pub fn set_border_color(&self, color: Color) -> ConditionalFormatDataBar {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ConditionalFormatDataBar::new());
        inner = inner.set_border_color(xlsx::Color::from(color));
        *lock = inner;
        ConditionalFormatDataBar {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the color of the fill for negative values in the data bar.
    ///
    /// Set the fill color for negative values in a data bar conditional format.
    /// By default the negative data bar color is `#FF0000` (red).
    ///
    /// # Parameters
    ///
    /// - `color`: The color property defined by a {@link Color} enum value or a
    ///   type that can convert `Into` a {@link Color}.
    #[wasm_bindgen(js_name = "setNegativeFillColor", skip_jsdoc)]
    pub fn set_negative_fill_color(&self, color: Color) -> ConditionalFormatDataBar {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ConditionalFormatDataBar::new());
        inner = inner.set_negative_fill_color(xlsx::Color::from(color));
        *lock = inner;
        ConditionalFormatDataBar {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the color of the border for negative values in the data bar.
    ///
    /// Set the border color for negative values in a data bar conditional
    /// format. By default the border is the same color as the data bar:
    /// is `#FF0000` (red).
    ///
    /// # Parameters
    ///
    /// - `color`: The color property defined by a {@link Color} enum value or a
    ///   type that can convert `Into` a {@link Color}.
    #[wasm_bindgen(js_name = "setNegativeBorderColor", skip_jsdoc)]
    pub fn set_negative_border_color(&self, color: Color) -> ConditionalFormatDataBar {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ConditionalFormatDataBar::new());
        inner = inner.set_negative_border_color(xlsx::Color::from(color));
        *lock = inner;
        ConditionalFormatDataBar {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the data bar fill to solid.
    ///
    /// By default Excel uses a gradient fill for data bar conditional formats.
    /// This option can be used to turn on a solid fill.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setSolidFill", skip_jsdoc)]
    pub fn set_solid_fill(&self, enable: bool) -> ConditionalFormatDataBar {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ConditionalFormatDataBar::new());
        inner = inner.set_solid_fill(enable);
        *lock = inner;
        ConditionalFormatDataBar {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Turn off the border for a data bar conditional format.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is on by default.
    #[wasm_bindgen(js_name = "setBorderOff", skip_jsdoc)]
    pub fn set_border_off(&self, enable: bool) -> ConditionalFormatDataBar {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ConditionalFormatDataBar::new());
        inner = inner.set_border_off(enable);
        *lock = inner;
        ConditionalFormatDataBar {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the direction of the data bar conditional format.
    ///
    /// Set the data bar conditional format direction to "Right to left", "Left
    /// to right" or "Context" (the default).
    ///
    /// # Parameters
    ///
    /// - `direction`: A {@link ConditionalFormatDataBarDirection} enum value.
    #[wasm_bindgen(js_name = "setDirection", skip_jsdoc)]
    pub fn set_direction(
        &self,
        direction: ConditionalFormatDataBarDirection,
    ) -> ConditionalFormatDataBar {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ConditionalFormatDataBar::new());
        inner = inner.set_direction(xlsx::ConditionalFormatDataBarDirection::from(direction));
        *lock = inner;
        ConditionalFormatDataBar {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Hide the values and only show the data bars.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setBarOnly", skip_jsdoc)]
    pub fn set_bar_only(&self, enable: bool) -> ConditionalFormatDataBar {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ConditionalFormatDataBar::new());
        inner = inner.set_bar_only(enable);
        *lock = inner;
        ConditionalFormatDataBar {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the position of the axis in a data bar.
    ///
    /// The position can be set to midpoint or turned off.
    ///
    /// # Parameters
    ///
    /// - `position`: A {@link ConditionalFormatDataBarAxisPosition} enum value.
    #[wasm_bindgen(js_name = "setAxisPosition", skip_jsdoc)]
    pub fn set_axis_position(
        &self,
        position: ConditionalFormatDataBarAxisPosition,
    ) -> ConditionalFormatDataBar {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ConditionalFormatDataBar::new());
        inner = inner.set_axis_position(xlsx::ConditionalFormatDataBarAxisPosition::from(position));
        *lock = inner;
        ConditionalFormatDataBar {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the color of the axis in the data bar.
    ///
    /// Set the axis color for a data bar conditional format. By default the
    /// axis color is `#000000` (black).
    ///
    /// # Parameters
    ///
    /// - `color`: The color property defined by a {@link Color} enum value or a
    ///   type that can convert `Into` a {@link Color}.
    #[wasm_bindgen(js_name = "setAxisColor", skip_jsdoc)]
    pub fn set_axis_color(&self, color: Color) -> ConditionalFormatDataBar {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ConditionalFormatDataBar::new());
        inner = inner.set_axis_color(xlsx::Color::from(color));
        *lock = inner;
        ConditionalFormatDataBar {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the data bar format to the original Excel 2007 style.
    ///
    /// The original Excel 2007 style was simpler than the post Excel 2010 style
    /// and had very limited configuration options.
    ///
    /// This is implemented for backward compatibility and for testing but is
    /// unlikely be be required by the end user.
    #[wasm_bindgen(js_name = "useClassicStyle", skip_jsdoc)]
    pub fn use_classic_style(&self) -> ConditionalFormatDataBar {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ConditionalFormatDataBar::new());
        inner = inner.use_classic_style();
        *lock = inner;
        ConditionalFormatDataBar {
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
    pub fn set_multi_range(&self, range: &str) -> ConditionalFormatDataBar {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ConditionalFormatDataBar::new());
        inner = inner.set_multi_range(range);
        *lock = inner;
        ConditionalFormatDataBar {
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
    pub fn set_stop_if_true(&self, enable: bool) -> ConditionalFormatDataBar {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ConditionalFormatDataBar::new());
        inner = inner.set_stop_if_true(enable);
        *lock = inner;
        ConditionalFormatDataBar {
            inner: Arc::clone(&self.inner),
        }
    }
}
