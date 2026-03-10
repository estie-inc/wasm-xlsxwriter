use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `Url` struct is used to define a worksheet URL.
///
/// The `Url` struct creates a URL type that can be used to write worksheet
/// URLs.
///
/// In general, you would use the
/// {@link Worksheet#writeUrl} with a string
/// representation of the URL, like this:
///
/// The URL will then be displayed as expected in Excel:
///
/// To differentiate a URL from an ordinary string (for example, when storing it
/// in a data structure), you can also represent the URL with a {@link Url} struct:
///
/// Using a `Url` struct also allows you to write a URL using the generic
/// {@link Worksheet#write} method:
///
/// There are three types of URLs/links supported by Excel and the
/// `rust_xlsxwriter` library:
///
/// 1. Web-based URIs like:
///
///    * `http://`, `https://`, `ftp://`, `ftps://`, `mailto:` and custom user
///      links like `protocol://domain`.
///
/// 2. Local file links using the `file://` URI.
///
///    * `file:///Book2.xlsx`
///    * `file:///..\Sales\Book2.xlsx`
///    * `file:///C:\Temp\Book1.xlsx`
///    * `file:///Book2.xlsx#Sheet1!A1`
///    * `file:///Book2.xlsx#'Sales Data'!A1:G5`
///
///    Most paths will be relative to the root folder, following the Windows
///    convention, so most paths should start with `file:///`. For links to
///    other Excel files, the URL string can include a sheet and cell reference
///    after the `"#"` anchor, as shown in the last two examples above. When
///    using Windows paths, like in the examples above, it is best to use a Rust
///    raw string to avoid issues with the backslashes:
///    `r"file:///C:\Temp\Book1.xlsx"`.
///
/// 3. Internal links to a cell or range of cells in the workbook using the
///    pseudo-URI `internal:`:
///
///    * `internal:Sheet2!A1`
///    * `internal:Sheet2!A1:G5`
///    * `internal:'Sales Data'!A1`
///
///    Worksheet references are typically of the form `Sheet1!A1`, where a
///    worksheet and target cell should be specified. You can also link to a
///    worksheet range using the standard Excel range notation like
///    `Sheet1!A1:B2`. Excel requires that worksheet names containing spaces or
///    non-alphanumeric characters are single-quoted as follows: `'Sales
///    Data'!A1`.
///
/// The library will escape the following characters in URLs as required by
/// Excel: ``\s " < > \ [ ] ` ^ { }``, unless the URL already contains `%xx`
/// style escapes. In that case, it is assumed that the URL was escaped
/// correctly by the user and will be passed directly to Excel.
///
/// Excel has a limit of around 2080 characters in the URL string. URLs beyond
/// this limit will raise an error when written.
#[derive(Clone)]
#[wasm_bindgen]
pub struct Url {
    pub(crate) inner: Arc<Mutex<xlsx::Url>>,
}

#[wasm_bindgen]
impl Url {
    #[wasm_bindgen(constructor)]
    pub fn new(link: &str) -> Url {
        Url {
            inner: Arc::new(Mutex::new(xlsx::Url::new(link))),
        }
    }
    /// Set the alternative text for the URL.
    ///
    /// Set an alternative, user friendly, text for the URL.
    ///
    /// # Parameters
    ///
    /// `text` - The alternative text, as a string or string like type.
    #[wasm_bindgen(js_name = "setText", skip_jsdoc)]
    pub fn set_text(&self, text: &str) -> Url {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Url::new(""));
        inner = inner.set_text(text);
        *lock = inner;
        Url {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the screen tip for the URL.
    ///
    /// Set a screen tip when the user does a mouseover of the URL.
    ///
    /// # Parameters
    ///
    /// `tip` - The URL tip, as a string or string like type.
    #[wasm_bindgen(js_name = "setTip", skip_jsdoc)]
    pub fn set_tip(&self, tip: &str) -> Url {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Url::new(""));
        inner = inner.set_tip(tip);
        *lock = inner;
        Url {
            inner: Arc::clone(&self.inner),
        }
    }
}
