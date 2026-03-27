use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `FilterCondition` struct is used to define autofilter rules.
///
/// Autofilter rules are associated with ranges created using
/// {@link autofilter}).
///
/// Excel supports two main types of filter conditions. The first, and most
/// common, is a list filter where the user selects the items to filter from a
/// list of all the values in the column range:
///
/// The other main type of filter is a custom filter where the user can specify
/// 1 or 2 conditions like ">= 4000" and "<= 6000":
///
/// In Excel these are mutually exclusive and you will need to choose one or the
/// other via the {@link FilterCondition}
/// {@link add_list_filter}(FilterCondition::add_list_filter) and
/// {@link add_custom_filter}(FilterCondition::add_custom_filter) methods.
#[derive(Clone)]
#[wasm_bindgen]
pub struct FilterCondition {
    pub(crate) inner: Arc<Mutex<xlsx::FilterCondition>>,
}

#[wasm_bindgen]
impl FilterCondition {
    #[wasm_bindgen(constructor)]
    pub fn new() -> FilterCondition {
        FilterCondition {
            inner: Arc::new(Mutex::new(xlsx::FilterCondition::new())),
        }
    }
    #[doc = r" Create a deep clone of this object."]
    #[wasm_bindgen(js_name = "clone")]
    pub fn deep_clone(&self) -> FilterCondition {
        FilterCondition {
            inner: Arc::new(Mutex::new(self.inner.lock().unwrap().clone())),
        }
    }
    /// Add a list filter to filter on blanks.
    ///
    /// Add a filter condition to a list filter to show blank cells. For
    /// autofilters Excel treats empty or whitespace only cells as "Blank".
    ///
    /// Filtering non-blanks can be done in two ways. See the second example
    /// below.
    #[wasm_bindgen(js_name = "addListBlanksFilter", skip_jsdoc)]
    pub fn add_list_blanks_filter(&self) -> FilterCondition {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::FilterCondition::new());
        inner = inner.add_list_blanks_filter();
        *lock = inner;
        FilterCondition {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Add an "or" logical condition for two custom filters.
    ///
    /// When two conditions are specified, like the example above, the logical
    /// operator defaults to "and", as in Excel. However, you can use the
    /// {@link add_custom_boolean_or}(FilterCondition::add_custom_boolean_or) method
    /// to get an "or" logical condition.
    #[wasm_bindgen(js_name = "addCustomBooleanOr", skip_jsdoc)]
    pub fn add_custom_boolean_or(&self) -> FilterCondition {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::FilterCondition::new());
        inner = inner.add_custom_boolean_or();
        *lock = inner;
        FilterCondition {
            inner: Arc::clone(&self.inner),
        }
    }
}
