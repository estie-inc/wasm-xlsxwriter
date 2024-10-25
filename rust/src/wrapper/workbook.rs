use std::sync::{Arc, Mutex};

use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

use crate::{
    error::XlsxError,
    wrapper::{doc_properties::DocProperties, map_xlsx_error, worksheet::Worksheet},
};

use super::WasmResult;

/// The `Workbook` struct represents an Excel file in its entirety. It is the
/// starting point for creating a new Excel xlsx file.
///
// TODO: example omitted
#[wasm_bindgen]
pub struct Workbook {
    inner: Arc<Mutex<xlsx::Workbook>>,
    next_sheet_index: usize,
}

#[wasm_bindgen]
impl Workbook {
    /// Create a new Workbook object to represent an Excel spreadsheet file.
    ///
    /// This constructor is used to create a new Excel workbook
    /// object. This is used to create worksheets and add data prior to saving
    /// everything to an xlsx file with {@link Workbook#saveToBufferSync}.
    ///
    /// **Note**: `rust_xlsxwriter` can only create new files. It cannot read or
    /// modify existing files.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(constructor)]
    pub fn new() -> Self {
        Workbook {
            inner: Arc::new(Mutex::new(xlsx::Workbook::new())),
            next_sheet_index: 0,
        }
    }

    /// Add a new worksheet to a workbook.
    ///
    /// The `addWorksheet()` method adds a new {{@link Worksheet} to a
    /// workbook.
    ///
    /// The worksheets will be given standard Excel name like `Sheet1`,
    /// `Sheet2`, etc. Alternatively, the name can be set using
    /// `worksheet.setName()`, see the example below and the docs for
    /// {@link Worksheet#setName}.
    ///
    /// The `addWorksheet()` method returns a borrowed mutable reference to a
    /// Worksheet instance owned by the Workbook so only one worksheet can be in
    /// existence at a time, see the example below. This limitation can be
    /// avoided, if necessary, by creating standalone Worksheet objects via
    /// {@link Worksheet::constructor} and then later adding them to the workbook with
    /// {@link Workbook#pushWorksheet}.
    ///
    /// See also the documentation on [Creating worksheets] and working with the
    /// borrow checker.
    ///
    /// [Creating worksheets]: ../worksheet/index.html#creating-worksheets
    ///
    /// @returns {Worksheet} - The worksheet object.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "addWorksheet", skip_jsdoc)]
    pub fn add_worksheet(&mut self) -> Worksheet {
        let index = self.next_sheet_index;
        self.next_sheet_index += 1;
        let mut workbook = self.inner.lock().unwrap();
        let _ = workbook.add_worksheet();
        Worksheet {
            workbook: Arc::clone(&self.inner),
            index,
        }
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
    /// using `worksheetFromIndex()`. The standard borrow checking rules still
    /// apply so you will have to give up ownership of any other worksheet
    /// reference prior to calling this method. See the example below.
    ///
    /// See also {@link Workbook#worksheetFromName} and the documentation on
    /// [Creating worksheets] and working with the borrow checker.
    ///
    /// [Creating worksheets]: ../worksheet/index.html#creating-worksheets
    ///
    /// @param {number} index - The index of the worksheet to get a reference to.
    /// @returns {Worksheet} - The worksheet object.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::UnknownWorksheetNameOrIndex`] - Error when trying to
    ///   retrieve a worksheet reference by index. This is usually an index out
    ///   of bounds error.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "worksheetFromIndex", skip_jsdoc)]
    pub fn worksheet_from_index(&self, index: usize) -> WasmResult<Worksheet> {
        // Reimplementation of [`rust_xlsxwriter::Workbook::worksheet_from_name()`]
        let mut workbook = self.inner.lock().unwrap();
        let _ = map_xlsx_error(workbook.worksheet_from_index(index))?;
        Ok(Worksheet {
            workbook: Arc::clone(&self.inner),
            index,
        })
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
    /// using `worksheetFromName()`. The standard borrow checking rules still
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
    /// @param {string} name - The name of the worksheet to get a reference to.
    /// @returns {Worksheet} - The worksheet object.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::UnknownWorksheetNameOrIndex`] - Error when trying to
    ///   retrieve a worksheet reference by index. This is usually an index out
    ///   of bounds error.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "worksheetFromName", skip_jsdoc)]
    pub fn worksheet_from_name(&self, name: &str) -> WasmResult<Worksheet> {
        // Reimplementation of [`rust_xlsxwriter::Workbook::worksheet_from_name()`]
        let mut workbook = self.inner.lock().unwrap();
        for (index, worksheet) in workbook.worksheets().iter().enumerate() {
            if worksheet.name() == name {
                return Ok(Worksheet {
                    workbook: Arc::clone(&self.inner),
                    index,
                });
            }
        }
        Err(XlsxError::Xlsx(
            xlsx::XlsxError::UnknownWorksheetNameOrIndex(name.to_string()),
        ))
    }

    /// Create a defined name in the workbook to use as a variable.
    ///
    /// The `defineName()` method is used to defined a variable name that can
    /// be used to represent a value, a single cell or a range of cells in a
    /// workbook. These are sometimes referred to as a "Named Ranges".
    ///
    /// Defined names are generally used to simplify or clarify formulas by
    /// using descriptive variable names. For example:
    ///
    /// ```text
    ///     // Global workbook name.
    ///     workbook.define_name("Exchange_rate", "=0.96")?;
    ///     worksheet.write_formula(0, 0, "=Exchange_rate")?;
    /// ```
    ///
    /// A name defined like this is "global" to the workbook and can be used in
    /// any worksheet in the workbook.  It is also possible to define a
    /// local/worksheet name by prefixing it with the sheet name using the
    /// syntax `"sheetname!defined_name"`:
    ///
    /// ```text
    ///     // Local worksheet name.
    ///     workbook.define_name('Sheet2!Sales', '=Sheet2!$G$1:$G$10')?;
    /// ```
    ///
    /// See the full example below.
    ///
    /// Note, Excel has limitations on names used in defined names. For example
    /// it must start with a letter or underscore and cannot contain a space or
    /// any of the characters: `,/*[]:\"'`. It also cannot look like an Excel
    /// range such as `A1`, `XFD12345` or `R1C1`. If in doubt it best to test
    /// the name in Excel first.
    ///
    /// For local defined names sheet name must exist (at the time of saving)
    /// and if the sheet name contains spaces or special characters you must
    /// follow the Excel convention and enclose it in single quotes:
    ///
    /// ```text
    ///     workbook.define_name("'New Data'!Sales", ""=Sheet2!$G$1:$G$10")?;
    /// ```
    ///
    /// The rules for names in Excel are explained in the Microsoft Office
    /// documentation on how to [Define and use names in
    /// formulas](https://support.microsoft.com/en-us/office/define-and-use-names-in-formulas-4d0f13ac-53b7-422e-afd2-abd7ff379c64)
    /// and subsections.
    ///
    /// @param {string} name - The variable name to define.
    /// @param {string} formula - The formula, value or range that the name defines..
    ///
    /// # Errors
    ///
    /// - [`XlsxError::ParameterError`] - The following Excel error cases will
    ///   raise a `ParameterError` error:
    ///   * If the name doesn't start with a letter or underscore.
    ///   * If the name contains `,/*[]:\"'` or `space`.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "defineName", skip_jsdoc)]
    pub fn define_name(&mut self, name: &str, formula: &str) -> WasmResult<()> {
        let mut workbook = self.inner.lock().unwrap();
        let _ = workbook.define_name(name, formula);
        Ok(())
    }

    /// Save the Workbook as an xlsx file and return it as a byte vector.
    ///
    /// The workbook `saveToBufferSync()` returns the xlsx file as a
    /// `Vec<u8>` buffer suitable for streaming in a web application.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::SheetnameReused`] - Worksheet name is already in use in
    ///   the workbook.
    /// - [`XlsxError::IoError`] - A wrapper for various IO errors when creating
    ///   the xlsx file, or its sub-files.
    /// - [`XlsxError::ZipError`] - A wrapper for various zip errors when
    ///   creating the xlsx file, or its sub-files.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "saveToBufferSync")]
    pub fn save_to_buffer_sync(&self) -> WasmResult<Vec<u8>> {
        let mut workbook = self.inner.lock().unwrap();
        let buf = map_xlsx_error(workbook.save_to_buffer())?;
        Ok(buf)
    }

    /// Add a recommendation to open the file in “read-only” mode.
    ///
    /// This method can be used to set the Excel “Read-only Recommended” option
    /// that is available when saving a file. This presents the user of the file
    /// with an option to open it in "read-only" mode. This means that any
    /// changes to the file can’t be saved back to the same file and must be
    /// saved to a new file.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "readOnlyRecommended")]
    pub fn read_only_recommended(&self) {
        let mut workbook = self.inner.lock().unwrap();
        workbook.read_only_recommended();
    }

    /// Set the Excel document metadata properties.
    ///
    /// Set various Excel document metadata properties such as Author or
    /// Creation Date. It is used in conjunction with the {@link DocProperties}
    /// struct.
    ///
    /// @param {DocProperties} properties - A reference to a {@link DocProperties} object.
    #[wasm_bindgen(js_name = "setProperties", skip_jsdoc)]
    pub fn set_properties(&self, properties: &DocProperties) {
        let mut workbook = self.inner.lock().unwrap();
        workbook.set_properties(&properties.inner);
    }
}
