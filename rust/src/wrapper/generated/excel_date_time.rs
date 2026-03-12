use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ExcelDateTime` struct is used to represent an Excel date and/or time.
///
/// The `rust_xlsxwriter` library supports converting different date and time
/// types to an Excel serial datetime. The main option is the inbuilt
/// {@link ExcelDateTime} type which has a limited but workable set of conversion
/// methods and which only targets Excel specific dates and times. The second is
/// via the optional external crates {@link Chrono} and {@link Jiff} which have more
/// comprehensive functionality for dealing with dates and times.
///
/// {@link Jiff}: https://docs.rs/chrono/latest/jiff
/// {@link Chrono}: https://docs.rs/chrono/latest/chrono
///
/// Here is an example using `ExcelDateTime` to write some dates and times:
///
/// Output file:
///
/// ## Datetimes in Excel
///
/// Datetimes in Excel are serial dates with days counted from an epoch (usually
/// 1900-01-01) and where the time is a percentage/decimal of the milliseconds
/// in the day. Both the date and time are stored in the same f64 value. For
/// example, 2023/01/01 12:00:00 is stored as 44927.5.
///
/// Datetimes in Excel must also be formatted with a number format like
/// `"yyyy/mm/dd hh:mm"` or otherwise they will appear as numbers (which
/// technically they are).
///
/// Excel doesn't use timezones or try to convert or encode timezone information
/// in any way so they aren't supported by `rust_xlsxwriter`.
///
/// Excel can also save dates in a text ISO 8601 format when the file is saved
/// using the "Strict Open XML Spreadsheet" option in the "Save" dialog. However
/// this is rarely used in practice and isn't supported by `rust_xlsxwriter`.
///
/// ## `chrono` and `jiff` vs. `ExcelDateTime`
///
/// The `rust_xlsxwriter` native `ExcelDateTime` type provides most of the
/// functionality that you will need to work with Excel dates and times.
///
/// For anything more advanced you can use the naive/civil date/time variants of
/// {@link Chrono} and {@link Jiff}, particularly if you are interacting with code that
/// already uses `Chrono` or `Jiff`.
///
/// The supported `Chrono` or `Jiff` date/time types are:
///
/// - {@link ExcelDateTime}: The inbuilt `rust_xlsxwriter` datetime type.
/// - {@link Chrono} naive types:
///   - {@link chrono#NaiveDateTime}
///   - {@link chrono#NaiveDate}
///   - {@link chrono#NaiveTime}
/// - {@link Jiff} civil types:
///   - {@link jiff#civil::Datetime}
///   - {@link jiff#civil::Date}
///   - {@link jiff#civil::Time}
///
/// {@link ExcelDateTime}: crate::ExcelDateTime
///
/// {@link Chrono}: https://docs.rs/chrono/latest/chrono
/// {@link chrono#NaiveDate}:
///     https://docs.rs/chrono/latest/chrono/naive/struct.NaiveDate.html
/// {@link chrono#NaiveTime}:
///     https://docs.rs/chrono/latest/chrono/naive/struct.NaiveTime.html
/// {@link chrono#NaiveDateTime}:
///     https://docs.rs/chrono/latest/chrono/naive/struct.NaiveDateTime.html
///
/// {@link Jiff}: https://docs.rs/jiff/latest/jiff
/// {@link jiff#civil::Datetime}:
///     https://docs.rs/jiff/latest/jiff/civil/struct.DateTime.html
/// {@link jiff#civil::Date}:
///     https://docs.rs/jiff/latest/jiff/civil/struct.Date.html
/// {@link jiff#civil::Time}:
///     https://docs.rs/jiff/latest/jiff/civil/struct.Time.html
///
/// All date/time APIs in `rust_xlsxwriter` support all of these types and the
/// `ExcelDateTime` method names are similar to `Chrono` method names to allow
/// easier portability between the two. Only the naive and civil types are
/// supported since Excel doesn't support timezones.
///
/// In order to use {@link Chrono} or {@link Jiff} with `rust_xlsxwriter` APIs you must
/// enable the optional `chrono` and `jiff` features when adding
/// `rust_xlsxwriter` to your `Cargo.toml`.
///
/// {@link Jiff}: https://docs.rs/jiff/latest/jiff
/// {@link Chrono}: https://docs.rs/chrono/latest/chrono
#[derive(Clone)]
#[wasm_bindgen]
pub struct ExcelDateTime {
    pub(crate) inner: Arc<Mutex<xlsx::ExcelDateTime>>,
}

#[wasm_bindgen]
impl ExcelDateTime {
    #[doc = r" Create a deep clone of this object."]
    #[wasm_bindgen(js_name = "clone")]
    pub fn deep_clone(&self) -> ExcelDateTime {
        ExcelDateTime {
            inner: Arc::new(Mutex::new(self.inner.lock().unwrap().clone())),
        }
    }
    /// Convert the `ExcelDateTime` to an Excel serial date.
    ///
    /// An Excel serial date is a f64 number that represents the time since the
    /// Excel epoch. This method is mainly used internally when converting an
    /// `ExcelDateTime` instance to an Excel datetime. The method is exposed
    /// publicly to allow some limited manipulation of the date/time in
    /// conjunction with
    /// {@link from_serial_datetime}(ExcelDateTime::from_serial_datetime).
    #[wasm_bindgen(js_name = "toExcel", skip_jsdoc)]
    pub fn to_excel(&self) -> f64 {
        let lock = self.inner.lock().unwrap();
        lock.to_excel()
    }
}
