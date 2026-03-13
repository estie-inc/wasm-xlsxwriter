use std::sync::Arc;

use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

use crate::wrapper::{FilterCondition, FilterCriteria};

#[wasm_bindgen]
impl FilterCondition {
    /// Add a value to the filter list. Accepts a number or string.
    #[wasm_bindgen(js_name = "addListFilter")]
    pub fn add_list_filter(&self, value: JsValue) -> FilterCondition {
        let mut lock = self.inner.lock().unwrap();
        let inner = std::mem::replace(&mut *lock, xlsx::FilterCondition::new());
        *lock = if let Some(n) = value.as_f64() {
            inner.add_list_filter(n)
        } else {
            let s = value.as_string().unwrap_or_default();
            inner.add_list_filter(s.as_str())
        };
        FilterCondition {
            inner: Arc::clone(&self.inner),
        }
    }

    /// Add a custom filter with criteria and value. Accepts a number or string.
    #[wasm_bindgen(js_name = "addCustomFilter")]
    pub fn add_custom_filter(
        &self,
        criteria: FilterCriteria,
        value: JsValue,
    ) -> FilterCondition {
        let mut lock = self.inner.lock().unwrap();
        let inner = std::mem::replace(&mut *lock, xlsx::FilterCondition::new());
        let xlsx_criteria = xlsx::FilterCriteria::from(criteria);
        *lock = if let Some(n) = value.as_f64() {
            inner.add_custom_filter(xlsx_criteria, n)
        } else {
            let s = value.as_string().unwrap_or_default();
            inner.add_custom_filter(xlsx_criteria, s.as_str())
        };
        FilterCondition {
            inner: Arc::clone(&self.inner),
        }
    }
}
