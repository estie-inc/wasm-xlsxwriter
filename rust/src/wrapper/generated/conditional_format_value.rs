use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ConditionalFormatValue` struct represents conditional format value
/// types.
///
/// Excel supports various types when specifying values in a conditional format
/// such as numbers, strings, dates, times and cell references.
/// `ConditionalFormatValue` is used to support a similar generic interface to
/// conditional format values. It supports:
///
/// - Numbers: Any Rust number that can convert `Into` `f64`.
/// - Strings: Any Rust string type that can convert into String such as
///   {@link &str}, `String`, `&String` and `Cow<'_, str>`.
/// - Booleans: Any Rust `bool` value.
/// - Dates/times: {@link ExcelDateTime} values. If the `chrono` feature is enabled
///   {@link chrono#NaiveDateTime}, {@link chrono#NaiveDate} and
///   {@link chrono#NaiveTime}. If the `jiff` feature is enabled
///   {@link jiff#civil::Datetime}, {@link jiff#civil::Date} and
///   {@link jiff#civil::Time}.
/// - Cell ranges: Use {@link Formula} in order to distinguish from strings. For
///   example `Formula::new(=A1)`.
///
/// {@link chrono#NaiveDate}: https://docs.rs/chrono/latest/chrono/naive/struct.NaiveDate.html
/// {@link chrono#NaiveTime}: https://docs.rs/chrono/latest/chrono/naive/struct.NaiveTime.html
/// {@link chrono#NaiveDateTime}: https://docs.rs/chrono/latest/impl
///
/// {@link jiff#civil::Date}: https://docs.rs/jiff/latest/jiff/civil/struct.Date.html
/// {@link jiff#civil::Time}: https://docs.rs/jiff/latest/jiff/civil/struct.Time.html
/// {@link jiff#civil::Datetime}: https://docs.rs/jiff/latest/jiff/civil/struct.DateTime.html
#[derive(Clone)]
#[wasm_bindgen]
pub struct ConditionalFormatValue {
    pub(crate) inner: Arc<Mutex<xlsx::ConditionalFormatValue>>,
}

#[wasm_bindgen]
impl ConditionalFormatValue {}
