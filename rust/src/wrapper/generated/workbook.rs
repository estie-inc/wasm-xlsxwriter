use crate::wrapper::DocProperties;
use crate::wrapper::Format;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `Workbook` struct represents an Excel file in its entirety. It is the
/// starting point for creating a new Excel xlsx file.
#[derive(Clone)]
#[wasm_bindgen]
pub struct Workbook {
    pub(crate) inner: Arc<Mutex<xlsx::Workbook>>,
}

#[wasm_bindgen]
impl Workbook {
    #[wasm_bindgen(constructor)]
    pub fn new() -> Workbook {
        Workbook {
            inner: Arc::new(Mutex::new(xlsx::Workbook::new())),
        }
    }
    /// Add a new chartsheet to a workbook.
    ///
    /// The `add_chartsheet()` method adds a new "chartsheet" {@link Worksheet} to a
    /// workbook.
    ///
    /// A Chartsheet in Excel is a specialized type of worksheet that doesn't
    /// have cells but instead is used to display a single chart. It supports
    /// worksheet display options such as headers and footers, margins, tab
    /// selection, and print properties.
    ///
    /// The chartsheets will be given standard Excel name like `Chart1`,
    /// `Chart2`, etc. Alternatively, the name can be set using
    /// {@link Worksheet#setName}.
    ///
    /// The `add_worksheet()` method returns a borrowed mutable reference to a
    /// Worksheet instance owned by the Workbook so only one worksheet can be in
    /// existence at a time. This limitation can be avoided, if necessary, by
    /// creating standalone Worksheet objects via {@link Worksheet#new} and then
    /// later adding them to the workbook with {@link Workbook#pushWorksheet}.
    ///
    /// See also the documentation on [Creating worksheets] and working with the
    /// borrow checker.
    ///
    /// [Creating worksheets]: ../worksheet/index.html#creating-worksheets
    #[wasm_bindgen(js_name = "addChartsheet", skip_jsdoc)]
    pub fn add_chartsheet(&self) -> &mut Worksheet {
        let lock = self.inner.lock().unwrap();
        lock.add_chartsheet()
    }
    /// Get a worksheet reference by index.
    ///
    /// Get a reference to a worksheet created via {@link Workbook#addWorksheet}
    /// using an index based on the creation order.
    ///
    /// Due to borrow checking rules you can only have one active reference to a
    /// worksheet object created by `add_worksheet()` since that method always
    /// returns a mutable reference. For a workbook with multiple worksheets
    /// this restriction is generally workable if you can create and use the
    /// worksheets sequentially since you will only need to have one reference
    /// at any one time. However, if you can't structure your code to work
    /// sequentially then you get a reference to a previously created worksheet
    /// using `worksheet_from_index()`. The standard borrow checking rules still
    /// apply so you will have to give up ownership of any other worksheet
    /// reference prior to calling this method. See the example below.
    ///
    /// See also {@link Workbook#worksheetFromName} and the documentation on
    /// [Creating worksheets] and working with the borrow checker.
    ///
    /// [Creating worksheets]: ../worksheet/index.html#creating-worksheets
    ///
    /// # Parameters
    ///
    /// - `index`: The index of the worksheet to get a reference to.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#UnknownWorksheetNameOrIndex} - Error when trying to
    ///   retrieve a worksheet reference by index. This is usually an index out
    ///   of bounds error.
    #[wasm_bindgen(js_name = "worksheetFromIndex", skip_jsdoc)]
    pub fn worksheet_from_index(&self, index: usize) -> Result<&mut Worksheet> {
        let lock = self.inner.lock().unwrap();
        lock.worksheet_from_index(index)
    }
    /// Get a worksheet reference by name.
    ///
    /// Get a reference to a worksheet created via {@link Workbook#addWorksheet}
    /// using the sheet name.
    ///
    /// Due to borrow checking rules you can only have one active reference to a
    /// worksheet object created by `add_worksheet()` since that method always
    /// returns a mutable reference. For a workbook with multiple worksheets
    /// this restriction is generally workable if you can create and use the
    /// worksheets sequentially since you will only need to have one reference
    /// at any one time. However, if you can't structure your code to work
    /// sequentially then you get a reference to a previously created worksheet
    /// using `worksheet_from_name()`. The standard borrow checking rules still
    /// apply so you will have to give up ownership of any other worksheet
    /// reference prior to calling this method. See the example below.
    ///
    /// Worksheet names are usually "Sheet1", "Sheet2", etc., or else a user
    /// define name that was set using {@link Worksheet#setName}. You can also
    /// use the {@link Worksheet#name} method to get the name.
    ///
    /// See also {@link Workbook#worksheetFromIndex} and the documentation on
    /// [Creating worksheets] and working with the borrow checker.
    ///
    /// [Creating worksheets]: ../worksheet/index.html#creating-worksheets
    ///
    /// # Parameters
    ///
    /// - `name`: The name of the worksheet to get a reference to.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#UnknownWorksheetNameOrIndex} - Error when trying to
    ///   retrieve a worksheet reference by index. This is usually an index out
    ///   of bounds error.
    #[wasm_bindgen(js_name = "worksheetFromName", skip_jsdoc)]
    pub fn worksheet_from_name(&self, sheetname: &str) -> Result<&mut Worksheet> {
        let lock = self.inner.lock().unwrap();
        lock.worksheet_from_name(sheetname)
    }
    /// Create a defined name in the workbook to use as a variable.
    ///
    /// The `define_name()` method is used to define a variable name that can
    /// be used to represent a value, a single cell, or a range of cells in a
    /// workbook. These are sometimes referred to as "Named Ranges."
    ///
    /// Defined names are generally used to simplify or clarify formulas by
    /// using descriptive variable names. For example:
    ///
    /// A name defined like this is "global" to the workbook and can be used in
    /// any worksheet in the workbook.  It is also possible to define a
    /// local/worksheet name by prefixing it with the sheet name using the
    /// syntax `"sheetname!defined_name"`:
    ///
    /// See the full example below.
    ///
    /// Note, Excel has limitations on names used in defined names. For example,
    /// it must start with a letter or underscore and cannot contain a space or
    /// any of the characters: `,/*[]:\"'`. It also cannot look like an Excel
    /// range such as `A1`, `XFD12345`, or `R1C1`. If in doubt, it is best to test
    /// the name in Excel first.
    ///
    /// For local defined names sheet name must exist (at the time of saving)
    /// and if the sheet name contains spaces or special characters you must
    /// follow the Excel convention and enclose it in single quotes:
    ///
    /// The rules for names in Excel are explained in the Microsoft Office
    /// documentation on how to [Define and use names in
    /// formulas](https://support.microsoft.com/en-us/office/define-and-use-names-in-formulas-4d0f13ac-53b7-422e-afd2-abd7ff379c64)
    /// and subsections.
    ///
    /// # Parameters
    ///
    /// - `name`: The variable name to define.
    /// - `formula`: The formula, value or range that the name defines..
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#ParameterError} - The following Excel error cases will
    ///   raise a `ParameterError` error:
    ///   * If the name doesn't start with a letter or underscore.
    ///   * If the name contains `,/*[]:\"'` or `space`.
    #[wasm_bindgen(js_name = "defineName", skip_jsdoc)]
    pub fn define_name(&self, name: &str, formula: &str) -> WasmResult<Workbook> {
        let mut lock = self.inner.lock().unwrap();
        lock.define_name(name, formula)?;
        Ok(Workbook {
            inner: Arc::clone(&self.inner),
        })
    }
    /// Set the Excel document metadata properties.
    ///
    /// Set various Excel document metadata properties such as Author or
    /// Creation Date. It is used in conjunction with the {@link DocProperties}
    /// struct.
    ///
    /// # Parameters
    ///
    /// - `properties`: A reference to a {@link DocProperties} object.
    #[wasm_bindgen(js_name = "setProperties", skip_jsdoc)]
    pub fn set_properties(&self, properties: DocProperties) -> Workbook {
        let mut lock = self.inner.lock().unwrap();
        lock.set_properties(&properties.inner);
        Workbook {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the default cell format for a workbook.
    ///
    /// The `rust_xlsxwriter` library uses the Excel 2007 default cell format of
    /// "Calibri 11" for all worksheets. If required, it is possible to change
    /// the default format, mainly the font properties, to something like the
    /// more recent Excel default of "Aptos Narrow 11", the older "Arial 10", or
    /// an international font such as "ＭＳ Ｐゴシック".
    ///
    /// Changing the default format, and font, changes the default column width
    /// and row height for a worksheet and as a result it changes the
    /// positioning and scaling of objects such as images and charts. As such,
    /// it is also necessary to set the default column pixel width and row pixel
    /// height when changing the default format/font. These dimensions can be
    /// obtained by clicking on the row and column gridlines in a sample
    /// worksheet in Excel. See the example below for a "Aptos Narrow 11"
    /// workbook where the required height and width dimensions would be 20 and
    /// 64:
    ///
    /// src="https://rustxlsxwriter.github.io/images/workbook_aptos_narrow.png">
    ///
    /// These dimensions should be obtained from a Windows version of Excel
    /// since the macOS versions use non-standard dimensions.
    ///
    /// Only fonts that have the following column pixel width are currently
    /// supported: 56, 64, 72, 80, 96, 104 and 120. However, this should cover
    /// the most common fonts that Excel uses. If you need to support a
    /// different font/width combination please open a feature request in the
    /// GitHub repository with a sample blank workbook.
    ///
    /// This method must be called before adding any worksheets to the workbook
    /// so that the default format can be shared with new worksheets.
    ///
    /// This method doesn't currently adapt the {@link Worksheet#autofit} method
    /// for the new format font but this will hopefully be added in a future
    /// release.
    ///
    /// # Parameters
    ///
    /// - `format`: The new default {@link Format} property for the workbook.
    /// - `row_height`: The default row height in pixels.
    /// - `col_width`: The default column width in pixels.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#DefaultFormatError} - This error occurs if you try to
    ///   set the default format after a worksheet has been added to the
    ///   workbook, or if the pixel column width is not one of the supported
    ///   values shown above.
    #[wasm_bindgen(js_name = "setDefaultFormat", skip_jsdoc)]
    pub fn set_default_format(
        &self,
        format: Format,
        row_height: u32,
        col_width: u32,
    ) -> WasmResult<Workbook> {
        let mut lock = self.inner.lock().unwrap();
        lock.set_default_format(&format.inner, row_height, col_width)?;
        Ok(Workbook {
            inner: Arc::clone(&self.inner),
        })
    }
    /// Change the default workbook theme to the Excel 2023 Office/Aptos theme.
    ///
    /// Excel uses themes to define default fonts and colors for a workbook. The
    /// default theme in Excel from 2007 to 2022 was the "Office" theme with the
    /// Calibri 11 font and an associated color palette. In Excel 2023 and later
    /// the "Office" theme uses the Aptos 11 font and a different color palette.
    /// The older "Office" theme is now referred to as the "Office 2013-2022"
    /// theme.
    ///
    /// The `rust_xlsxwriter` library uses the original "Office" theme with
    /// Calibri 11 as the default font but, if required,
    /// `use_excel_2023_theme()` can be used to change to the newer Excel 2023
    /// "Office" theme with Aptos Narrow 11.
    ///
    /// Changing the theme and default font also changes the default row height
    /// and column width which in turn affects the positioning of objects such
    /// as images and charts. These changes are handled automatically.
    ///
    /// This method must be called before adding any worksheets to the workbook
    /// so that the default format can be shared with new worksheets.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#DefaultFormatError} - This error occurs if you try to
    ///   change the default theme, and by extension the default format, after a
    ///   worksheet has been added to the workbook.
    #[wasm_bindgen(js_name = "useExcel2023Theme", skip_jsdoc)]
    pub fn use_excel_2023_theme(&self) -> WasmResult<Workbook> {
        let mut lock = self.inner.lock().unwrap();
        lock.use_excel_2023_theme()?;
        Ok(Workbook {
            inner: Arc::clone(&self.inner),
        })
    }
    /// Use zip large file/ZIP64 extensions.
    ///
    /// The `rust_xlsxwriter` library uses the [zip.rs] crate to provide the zip
    /// container for the xlsx file that it generates. The size limit for a
    /// standard zip file is 4GB for the overall container or for any of the
    /// uncompressed files within it.  Anything greater than that requires
    /// [ZIP64] support. In practice this would apply to worksheets with
    /// approximately 150 million cells, or more.
    ///
    /// The `use_zip_large_file()` option enables ZIP64/large file support by
    /// enabling the `zip.rs` {@link large_file} option. Here is what the `zip.rs`
    /// library says about the `large_file()` option:
    ///
    /// > If `large_file()` is set to false and the file exceeds the limit, an
    /// > I/O error is thrown and the file is aborted. If set to true, readers
    /// > will require ZIP64 support and if the file does not exceed the limit,
    /// > 20 B are wasted. The default is false.
    ///
    /// You can interpret this to mean that it is safe/efficient to turn on
    /// large file mode by default if you anticipate that your application may
    /// generate files that exceed the 4GB limit. At least for Excel. Other
    /// applications may have issues if they don't support ZIP64 extensions.
    ///
    /// [zip.rs]: https://crates.io/crates/zip
    /// [ZIP64]: https://en.wikipedia.org/wiki/ZIP_(file_format)#ZIP64
    /// {@link large_file}:
    ///     https://docs.rs/zip/latest/zip/write/type.SimpleFileOptions.html#method.large_file
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "useZipLargeFile", skip_jsdoc)]
    pub fn use_zip_large_file(&self, enable: bool) -> Workbook {
        let mut lock = self.inner.lock().unwrap();
        lock.use_zip_large_file(enable);
        Workbook {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the workbook name used in VBA macros.
    ///
    /// This method can be used to set the VBA name for the workbook. This is
    /// sometimes required when a VBA macro included via
    /// {@link Workbook#addVbaProject} makes reference to the workbook with a
    /// name other than the default Excel VBA name of `ThisWorkbook`.
    ///
    /// See also the
    /// {@link Worksheet#setVbaName}) method
    /// for setting a worksheet VBA name.
    ///
    /// The name must be a valid Excel VBA object name as defined by the
    /// following rules:
    ///
    /// - The name must be less than 32 characters.
    /// - The name can only contain word characters: letters, numbers and
    ///   underscores.
    /// - The name must start with a letter.
    /// - The name cannot be blank.
    ///
    /// # Parameters
    ///
    /// - `name`: The vba name. It must follow the Excel rules, shown above.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#VbaNameError} - The name doesn't meet one of Excel's
    ///   criteria, shown above.
    #[wasm_bindgen(js_name = "setVbaName", skip_jsdoc)]
    pub fn set_vba_name(&self, name: &str) -> WasmResult<Workbook> {
        let mut lock = self.inner.lock().unwrap();
        lock.set_vba_name(name)?;
        Ok(Workbook {
            inner: Arc::clone(&self.inner),
        })
    }
    /// Add a recommendation to open the file in “read-only” mode.
    ///
    /// This method can be used to set the Excel “Read-only Recommended” option
    /// that is available when saving a file. This presents the user of the file
    /// with an option to open it in "read-only" mode. This means that any
    /// changes to the file can’t be saved back to the same file and must be
    /// saved to a new file.
    #[wasm_bindgen(js_name = "readOnlyRecommended", skip_jsdoc)]
    pub fn read_only_recommended(&self) -> Workbook {
        let mut lock = self.inner.lock().unwrap();
        lock.read_only_recommended();
        Workbook {
            inner: Arc::clone(&self.inner),
        }
    }
}
