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
}
