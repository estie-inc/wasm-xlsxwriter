use rust_xlsxwriter::{self as xlsx};
use wasm_bindgen::prelude::*;

/// The `Url` struct is used to define a worksheet url.
///
/// The `Url` struct creates a url type that can be used to write worksheet
/// urls.
///
/// In general you would use the
/// {@link Worksheet#writeUrl} with a string
/// representation of the url, like this:
///
/// TODO: example omitted
///
/// The url will then be displayed as expected in Excel:
///
/// <img src="https://rustxlsxwriter.github.io/images/url_intro1.png">
///
/// In order to differentiate a url from an ordinary string (for example when
/// storing it in a data structure) you can also represent the url with a
/// {@link Url} struct:
///
/// TODO: example omitted
///
/// Using a `Url` struct also allows you to write a url using the generic
/// {@link Worksheet#write} method:
///
/// TODO: example omitted
///
/// There are 3 types of url/link supported by Excel and the `rust_xlsxwriter`
/// library:
///
/// 1. Web based URIs like:
///
///    * `http://`, `https://`, `ftp://`, `ftps://` and `mailto:`.
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
///    other Excel files the url string can include a sheet and cell reference
///    after the `"#"` anchor, as shown in the last 2 examples above. When using
///    Windows paths, like in the examples above, it is best to use a Rust raw
///    string to avoid issues with the backslashes:
///    `r"file:///C:\Temp\Book1.xlsx"`.
///
/// 3. Internal links to a cell or range of cells in the workbook using the
///    pseudo-uri `internal:`:
///
///    * `internal:Sheet2!A1`
///    * `internal:Sheet2!A1:G5`
///    * `internal:'Sales Data'!A1`
///
///    Worksheet references are typically of the form `Sheet1!A1` where a
///    worksheet and target cell should be specified. You can also link to a
///    worksheet range using the standard Excel range notation like
///    `Sheet1!A1:B2`. Excel requires that worksheet names containing spaces or
///    non alphanumeric characters are single quoted as follows `'Sales
///    Data'!A1`.
///
/// The library will escape the following characters in URLs as required by
/// Excel, ``\s " < > \ [ ] ` ^ { }``, unless the URL already contains `%xx`
/// style escapes. In which case it is assumed that the URL was escaped
/// correctly by the user and will by passed directly to Excel.
///
/// Excel has a limit of around 2080 characters in the url string. Urls beyond
/// this limit will raise an error when written.
#[derive(Clone)]
#[wasm_bindgen]
pub struct Url {
    pub(crate) inner: xlsx::Url,
}

#[wasm_bindgen]
impl Url {
    /// Create a new Url struct.
    ///
    /// @param {string} link - A string like type representing a URL.
    /// @returns {Url} - The url object.
    #[wasm_bindgen(constructor, skip_jsdoc)]
    pub fn new(link: &str) -> Url {
        Url {
            inner: xlsx::Url::new(link),
        }
    }

    /// Set the alternative text for the url.
    ///
    /// Set an alternative, user friendly, text for the url.
    ///
    /// @param {string} text - The alternative text, as a string or string like type.
    /// @returns {Url} - The url object.
    #[wasm_bindgen(js_name = "setText", skip_jsdoc)]
    pub fn set_text(&self, text: &str) -> Url {
        Url {
            inner: self.clone().inner.set_text(text),
        }
    }

    /// Set the screen tip for the url.
    ///
    /// Set a screen tip when the user does a mouseover of the url.
    ///
    /// @param {string} tip - The url tip, as a string or string like type.
    /// @returns {Url} - The url object.
    #[wasm_bindgen(js_name = "setTip", skip_jsdoc)]
    pub fn set_tip(&self, tip: &str) -> Url {
        Url {
            inner: self.clone().inner.set_tip(tip),
        }
    }
}
