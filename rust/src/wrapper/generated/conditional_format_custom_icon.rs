use crate::wrapper::ConditionalFormatIconType;
use crate::wrapper::ConditionalFormatType;
use crate::wrapper::ConditionalFormatValue;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ConditionalFormatCustomIcon` struct represents an icon in an Icon Set
/// style conditional format.
///
/// The `ConditionalFormatCustomIcon` struct is create user defined icons for a
/// {@link ConditionalFormatIconSet} conditional format.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ConditionalFormatCustomIcon {
    pub(crate) inner: Arc<Mutex<xlsx::ConditionalFormatCustomIcon>>,
}

#[wasm_bindgen]
impl ConditionalFormatCustomIcon {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ConditionalFormatCustomIcon {
        ConditionalFormatCustomIcon {
            inner: Arc::new(Mutex::new(xlsx::ConditionalFormatCustomIcon::new())),
        }
    }
    #[doc = r" Create a deep clone of this object."]
    #[wasm_bindgen(js_name = "clone")]
    pub fn deep_clone(&self) -> ConditionalFormatCustomIcon {
        ConditionalFormatCustomIcon {
            inner: Arc::new(Mutex::new(self.inner.lock().unwrap().clone())),
        }
    }
    /// Set the rule for the custom icon.
    ///
    /// # Parameters
    ///
    /// - `rule_type`: A {@link ConditionalFormatType} enum value.
    /// - `value`: Any type that can convert into a {@link ConditionalFormatValue}
    ///   such as numbers, dates, times and formula ranges. String values are
    ///   ignored in this type of conditional format.
    #[wasm_bindgen(js_name = "setRule", skip_jsdoc)]
    pub fn set_rule(
        &self,
        rule_type: ConditionalFormatType,
        value: ConditionalFormatValue,
    ) -> ConditionalFormatCustomIcon {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ConditionalFormatCustomIcon::new());
        inner = inner.set_rule(
            xlsx::ConditionalFormatType::from(rule_type),
            value.inner.lock().unwrap().clone(),
        );
        *lock = inner;
        ConditionalFormatCustomIcon {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set a custom icon type.
    ///
    /// In Excel you can specify a custom icon for one of more icons in the
    /// default set using the following dialog:
    ///
    /// src="https://rustxlsxwriter.github.io/images/conditional_format_icon_custom_dialog.png">
    ///
    /// In `rust_xlsxwriter` you can emulate this by using the `set_icon_type()`
    /// API to specify the {@link ConditionalFormatIconType} and index to the icon
    /// within the icon type. For example the following are the
    /// {@link ConditionalFormatIconType#FiveBoxes} icons:
    ///
    /// So to use the fully filled box icon we would use index 4, see the
    /// example below.
    ///
    /// # Parameters
    ///
    /// - `icon_type`: A {@link ConditionalFormatIconType} enum value.
    /// - `index`: Index to the icon within the type. See the indexes shown in
    ///   the images for {@link ConditionalFormatIconType}.
    #[wasm_bindgen(js_name = "setIconType", skip_jsdoc)]
    pub fn set_icon_type(
        &self,
        icon_type: ConditionalFormatIconType,
        index: u8,
    ) -> ConditionalFormatCustomIcon {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ConditionalFormatCustomIcon::new());
        inner = inner.set_icon_type(xlsx::ConditionalFormatIconType::from(icon_type), index);
        *lock = inner;
        ConditionalFormatCustomIcon {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Turn off the icon in the cell.
    ///
    /// This is a variant of a custom icon setting:
    ///
    /// src="https://rustxlsxwriter.github.io/images/conditional_format_icon_custom_dialog.png">
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setNoIcon", skip_jsdoc)]
    pub fn set_no_icon(&self, enable: bool) -> ConditionalFormatCustomIcon {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ConditionalFormatCustomIcon::new());
        inner = inner.set_no_icon(enable);
        *lock = inner;
        ConditionalFormatCustomIcon {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the rule to be "greater than" instead of the Excel default of
    /// "greater than or equal to".
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setGreaterThan", skip_jsdoc)]
    pub fn set_greater_than(&self, enable: bool) -> ConditionalFormatCustomIcon {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::ConditionalFormatCustomIcon::new());
        inner = inner.set_greater_than(enable);
        *lock = inner;
        ConditionalFormatCustomIcon {
            inner: Arc::clone(&self.inner),
        }
    }
}
