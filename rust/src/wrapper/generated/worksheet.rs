use crate::wrapper::Color;
use crate::wrapper::Format;
use crate::wrapper::HeaderImagePosition;
use crate::wrapper::Image;
use crate::wrapper::ProtectionOptions;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone, Copy)]
pub enum WorksheetAccessor {
    AddWorksheet,
    AddChartsheet,
}

/// The `Worksheet` struct represents an Excel worksheet. It handles operations
/// such as writing data to cells or formatting the worksheet layout.
#[derive(Clone)]
#[wasm_bindgen]
pub struct Worksheet {
    pub(crate) parent: Arc<Mutex<xlsx::Workbook>>,
    pub(crate) accessor: WorksheetAccessor,
}

#[wasm_bindgen]
impl Worksheet {
    /// Set the worksheet name.
    ///
    /// Set the worksheet name. If no name is set the default Excel convention
    /// will be followed (Sheet1, Sheet2, etc.) in the order the worksheets are
    /// created.
    ///
    /// # Parameters
    ///
    /// - `name`: The worksheet name. It must follow the Excel rules, shown
    ///   below.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#SheetnameCannotBeBlank} - Worksheet name cannot be
    ///   blank.
    /// - {@link XlsxError#SheetnameLengthExceeded} - Worksheet name exceeds
    ///   Excel's limit of 31 characters.
    /// - {@link XlsxError#SheetnameContainsInvalidCharacter} - Worksheet name
    ///   cannot contain invalid characters: `[ ] : * ? / \`
    /// - {@link XlsxError#SheetnameStartsOrEndsWithApostrophe} - Worksheet name
    ///   cannot start or end with an apostrophe.
    #[wasm_bindgen(js_name = "setName", skip_jsdoc)]
    pub fn set_name(&self, name: &str) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_name(name),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_name(name),
        }?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
        })
    }
    /// Get the worksheet name.
    ///
    /// Get the worksheet name that was set automatically such as Sheet1,
    /// Sheet2, etc., or that was set by the user using
    /// {@link Worksheet#setName}.
    ///
    /// The worksheet name can be used to get a reference to a worksheet object
    /// using the
    /// {@link Workbook#worksheetFromName}
    /// method.
    #[wasm_bindgen(js_name = "name", skip_jsdoc)]
    pub fn name(&self) -> String {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().name(),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().name(),
        }
    }
    /// Insert a background image into a worksheet.
    ///
    /// A background image can be added to a worksheet to add a watermark or
    /// display a company logo. Excel repeats the image for the entirety of the
    /// worksheet.
    ///
    /// The image should be encapsulated in an {@link Image} object. See
    /// {@link Worksheet#insertImage} above for details on the supported image
    /// types.
    ///
    /// As an alternative to background images, it should be noted that the
    /// Microsoft Excel documentation recommends setting a watermark via an
    /// image in the worksheet header. An example of that technique is shown in
    /// the {@link Worksheet#setHeaderImage} examples.
    ///
    /// # Parameters
    ///
    /// - `image`: The {@link Image} to use as the worksheet background.
    #[wasm_bindgen(js_name = "insertBackgroundImage", skip_jsdoc)]
    pub fn insert_background_image(&self, image: Image) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock
                .add_worksheet()
                .insert_background_image(&*image.inner.lock().unwrap()),
            WorksheetAccessor::AddChartsheet => lock
                .add_chartsheet()
                .insert_background_image(&*image.inner.lock().unwrap()),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Make all worksheet notes visible when the file loads.
    ///
    /// By default Excel hides cell notes until the user mouses over the parent
    /// cell. However, if required you can make all worksheet notes visible when
    /// the worksheet loads. You can also make individual notes visible using
    /// the {@link Note#setVisible} method.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "showAllNotes", skip_jsdoc)]
    pub fn show_all_notes(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().show_all_notes(enable),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().show_all_notes(enable),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the default author name for all the notes in the worksheet.
    ///
    /// The Note author is the creator of the note. In Excel the author name is
    /// taken from the user name in the options/preference dialog. The note
    /// author name appears in two places: at the start of the note text in bold
    /// and at the bottom of the worksheet in the status bar.
    ///
    /// If no name is specified the default name "Author" will be applied to the
    /// note. The author name for individual notes can be set via the
    /// {@link Note#setAuthor} method. Alternatively
    /// this method can be used to set the default author name for all notes in
    /// a worksheet.
    ///
    /// # Parameters
    ///
    /// - `name`: The note author name. Must be less than or equal to the Excel
    ///   limit of 52 characters.
    #[wasm_bindgen(js_name = "setDefaultNoteAuthor", skip_jsdoc)]
    pub fn set_default_note_author(&self, name: &str) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_default_note_author(name),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_default_note_author(name),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Place the row outline group expand/collapse symbols above the range.
    ///
    /// This method toggles the Excel worksheet option to place the outline
    /// group expand/collapse symbols `[+]` and `[-]` above the group ranges
    /// instead of below for row ranges.
    ///
    /// In Excel this is a worksheet wide option and will apply to all row
    /// outlines in the worksheet.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "groupSymbolsAbove", skip_jsdoc)]
    pub fn group_symbols_above(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().group_symbols_above(enable),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().group_symbols_above(enable),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Place the column outline group expand/collapse symbols to the left of
    /// the range.
    ///
    /// This method toggles the Excel worksheet option to place the outline
    /// group expand/collapse symbols `[+]` and `[-]` to the left of the group
    /// ranges instead of to the right, for column ranges.
    ///
    /// In Excel this is a worksheet wide option and will apply to all column
    /// outlines in the worksheet.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "groupSymbolsToLeft", skip_jsdoc)]
    pub fn group_symbols_to_left(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().group_symbols_to_left(enable),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().group_symbols_to_left(enable),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the default row height for all rows in a worksheet, efficiently.
    ///
    /// This method can be used to efficiently set the default row height for
    /// all rows in a worksheet. It is efficient because it uses an Excel
    /// optimization to adjust the row heights with a single XML element. By
    /// contrast, using {@link Worksheet#setRowHeight} for every row in a
    /// worksheet would require a million XML elements and would result in a
    /// very large file.
    ///
    /// The height is specified in character units, where the default height is
    /// 15. Excel allows height values in increments of 0.25.
    ///
    /// Individual row heights can be set via {@link Worksheet#setRowHeight}.
    ///
    /// Note, there is no equivalent method for columns because the file format
    /// already optimizes the storage of a large number of contiguous columns.
    ///
    /// # Parameters
    ///
    /// - `height`: The row height in character units. Must be greater than 0.0.
    #[wasm_bindgen(js_name = "setDefaultRowHeight", skip_jsdoc)]
    pub fn set_default_row_height(&self, height: f64) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_default_row_height(height),
            WorksheetAccessor::AddChartsheet => {
                lock.add_chartsheet().set_default_row_height(height)
            }
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the default row height in pixels for all rows in a worksheet,
    /// efficiently.
    ///
    /// The `set_default_row_height_pixels()` method is used to change the
    /// default height of all rows in a worksheet. See
    /// {@link Worksheet#setDefaultRowHeight} above for an explanation.
    ///
    /// The height is specified in pixels, where the default height is 20.
    ///
    /// # Parameters
    ///
    /// - `height`: The row height in pixels.
    #[wasm_bindgen(js_name = "setDefaultRowHeightPixels", skip_jsdoc)]
    pub fn set_default_row_height_pixels(&self, height: u32) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => {
                lock.add_worksheet().set_default_row_height_pixels(height)
            }
            WorksheetAccessor::AddChartsheet => {
                lock.add_chartsheet().set_default_row_height_pixels(height)
            }
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Hide all unused rows in a worksheet, efficiently.
    ///
    /// This method can be used to efficiently hide unused rows in a worksheet.
    /// It is efficient because it uses an Excel optimization to hide the rows
    /// with a single XML element. By contrast, using
    /// {@link Worksheet#setRowHidden} for the majority of rows in a worksheet
    /// would require a million XML elements and would result in a very large
    /// file.
    ///
    /// "Unused" in this context means that the row doesn't contain data,
    /// formatting, or any changes such as the row height.
    ///
    /// To hide individual rows use the {@link Worksheet#setRowHidden} method.
    ///
    /// Note, there is no equivalent method for columns because the file format
    /// already optimizes the storage of a large number of contiguous columns.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "hideUnusedRows", skip_jsdoc)]
    pub fn hide_unused_rows(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().hide_unused_rows(enable),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().hide_unused_rows(enable),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the default cell format for a worksheet.
    ///
    /// See the {@link Workbook#setDefaultFormat} method for details.
    ///
    /// This method is only required when a worksheet is created independently
    /// of a workbook but you still wish to change the default format. The
    /// format must also be changed via {@link Workbook#setDefaultFormat} using
    /// the same format and dimensions in the parent workbook.
    ///
    /// {@link Workbook#setDefaultFormat}: crate::workbook::Workbook::set_default_format
    ///
    /// # Parameters
    ///
    /// - `format`: The new default {@link Format} property for the worksheet.
    /// - `row_height`: The default row height in pixels.
    /// - `col_width`: The default column width in pixels. Only fonts that have
    ///   the following column pixel width are supported: 56, 64, 72, 80, 96,
    ///   104 and 120.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#DefaultFormatError} - This error occurs if you try to
    ///   set the default format after formats have been added to a worksheet or
    ///   if the the pixel column width is one of the supported values shown
    ///   above.
    #[wasm_bindgen(js_name = "setDefaultFormat", skip_jsdoc)]
    pub fn set_default_format(
        &self,
        format: Format,
        row_height: u32,
        col_width: u32,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_default_format(
                &*format.inner.lock().unwrap(),
                row_height,
                col_width,
            ),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_default_format(
                &*format.inner.lock().unwrap(),
                row_height,
                col_width,
            ),
        }?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
        })
    }
    /// Turn off the option to automatically hide rows that don't match filters.
    ///
    /// Rows that don't match autofilter conditions are hidden by Excel at
    /// runtime. This feature isn't an automatic part of the file format and in
    /// practice it is necessary for the user to hide rows that don't match the
    /// applied filters. The `rust_xlsxwriter` library tries to do this
    /// automatically and in most cases will get it right, however, there may be
    /// cases where you need to manually hide some of the rows and may want to
    /// turn off the automatic handling using `filter_automatic_off()`.
    ///
    /// See [Auto-hiding filtered rows] in the User Guide.
    ///
    /// [Auto-hiding filtered rows]:
    ///     https://rustxlsxwriter.github.io/formulas/autofilters.html#auto-hiding-filtered-rows
    #[wasm_bindgen(js_name = "filterAutomaticOff", skip_jsdoc)]
    pub fn filter_automatic_off(&self) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().filter_automatic_off(),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().filter_automatic_off(),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Protect a worksheet from modification.
    ///
    /// The `protect()` method protects a worksheet from modification. It works
    /// by enabling a cell's `locked` and `hidden` properties, if they have been
    /// set. A **locked** cell cannot be edited and this property is on by
    /// default for all cells. A **hidden** cell will display the results of a
    /// formula but not the formula itself.
    ///
    /// src="https://rustxlsxwriter.github.io/images/protection_alert.png">
    ///
    /// These properties can be set using the
    /// {@link Format#setLocked}(Format::set_locked)
    /// {@link Format#setUnlocked}(Format::set_unlocked) and
    /// {@link Worksheet#setHidden}(Format::set_hidden) format methods. All cells
    /// have the `locked` property turned on by default (see the example below)
    /// so in general you don't have to explicitly turn it on.
    #[wasm_bindgen(js_name = "protect", skip_jsdoc)]
    pub fn protect(&self) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().protect(),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().protect(),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Protect a worksheet from modification with a password.
    ///
    /// The `protect_with_password()` method is like the
    /// {@link Worksheet#protect} method, see above, except that you can add an
    /// optional, weak, password to prevent modification.
    ///
    /// **Note**: Worksheet level passwords in Excel offer very weak protection.
    /// They do not encrypt your data and are very easy to deactivate. Full
    /// workbook encryption is not supported by `rust_xlsxwriter`. However, it
    /// is possible to encrypt an `rust_xlsxwriter` file using a third party
    /// open source tool called
    /// [msoffice-crypt](https://github.com/herumi/msoffice). This works for
    /// macOS, Linux and Windows:
    ///
    /// # Parameters
    ///
    /// - `password`: The password string. Note, only ascii text passwords are
    ///   supported. Passing the empty string "" is the same as turning on
    ///   protection without a password.
    #[wasm_bindgen(js_name = "protectWithPassword", skip_jsdoc)]
    pub fn protect_with_password(&self, password: &str) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().protect_with_password(password),
            WorksheetAccessor::AddChartsheet => {
                lock.add_chartsheet().protect_with_password(password)
            }
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Specify which worksheet elements should, or shouldn't, be protected.
    ///
    /// The `protect_with_password()` method is like the
    /// {@link Worksheet#protect} method, see above, except it also specifies
    /// which worksheet elements should, or shouldn't, be protected.
    ///
    /// You can specify which worksheet elements protection should be on or off
    /// via a {@link ProtectionOptions} struct reference. The Excel options with
    /// their default states are shown below:
    ///
    /// src="https://rustxlsxwriter.github.io/images/worksheet_protect_with_options1.png">
    ///
    /// # Parameters
    ///
    /// `options` - Worksheet protection options as defined by a
    /// {@link ProtectionOptions} struct reference.
    #[wasm_bindgen(js_name = "protectWithOptions", skip_jsdoc)]
    pub fn protect_with_options(&self, options: ProtectionOptions) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock
                .add_worksheet()
                .protect_with_options(&*options.inner.lock().unwrap()),
            WorksheetAccessor::AddChartsheet => lock
                .add_chartsheet()
                .protect_with_options(&*options.inner.lock().unwrap()),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Write the default formula result for worksheet formulas.
    ///
    /// The `rust_xlsxwriter` library doesn’t calculate the result of a formula
    /// written using {@link Worksheet#writeFormulaWithFormat} or
    /// {@link Worksheet#writeFormula}. Instead it stores the value 0 as the
    /// formula result. It then sets a global flag in the xlsx file to say that
    /// all formulas and functions should be recalculated when the file is
    /// opened.
    ///
    /// However, for `LibreOffice` the default formula result should be set to
    /// the empty string literal `""`, via the `set_formula_result_default()`
    /// method, to force calculation of the result.
    ///
    /// # Parameters
    ///
    /// - `result`: The default formula result to write to the cell.
    #[wasm_bindgen(js_name = "setFormulaResultDefault", skip_jsdoc)]
    pub fn set_formula_result_default(&self, result: &str) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => {
                lock.add_worksheet().set_formula_result_default(result)
            }
            WorksheetAccessor::AddChartsheet => {
                lock.add_chartsheet().set_formula_result_default(result)
            }
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Display the worksheet cells from right to left for some versions of
    /// Excel.
    ///
    /// The `set_right_to_left()` method is used to change the default direction
    /// of the worksheet from left-to-right, with the A1 cell in the top left,
    /// to right-to-left, with the A1 cell in the top right.
    ///
    /// This is useful when creating Arabic, Hebrew or other near or far eastern
    /// worksheets that use right-to-left as the default direction.
    ///
    /// Depending on your use case, and text, you may also need to use the
    /// {@link Format#setReadingDirection}
    /// method to set the direction of the text within the cells.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setRightToLeft", skip_jsdoc)]
    pub fn set_right_to_left(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_right_to_left(enable),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_right_to_left(enable),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Make a worksheet the active/initially visible worksheet in a workbook.
    ///
    /// The `set_active()` method is used to specify which worksheet is
    /// initially visible in a multi-sheet workbook. If no worksheet is set then
    /// the first worksheet is made the active worksheet, like in Excel.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setActive", skip_jsdoc)]
    pub fn set_active(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_active(enable),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_active(enable),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set a worksheet tab as selected.
    ///
    /// The `set_selected()` method is used to indicate that a worksheet is
    /// selected in a multi-sheet workbook.
    ///
    /// A selected worksheet has its tab highlighted. Selecting worksheets is a
    /// way of grouping them together so that, for example, several worksheets
    /// could be printed in one go. A worksheet that has been activated via the
    /// {@link Worksheet#setActive} method will also appear as selected.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setSelected", skip_jsdoc)]
    pub fn set_selected(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_selected(enable),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_selected(enable),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Hide a worksheet.
    ///
    /// The `set_hidden()` method is used to hide a worksheet. This can be used
    /// to hide a worksheet in order to avoid confusing a user with intermediate
    /// data or calculations.
    ///
    /// In Excel a hidden worksheet can not be activated or selected so this
    /// method is mutually exclusive with the {@link Worksheet#setActive} and
    /// {@link Worksheet#setSelected} methods. In addition, since the first
    /// worksheet will default to being the active worksheet, you cannot hide
    /// the first worksheet without activating another sheet.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setHidden", skip_jsdoc)]
    pub fn set_hidden(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_hidden(enable),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_hidden(enable),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Hide a worksheet. Can only be unhidden in Excel by VBA.
    ///
    /// The `set_very_hidden()` method can be used to hide a worksheet similar
    /// to the {@link Worksheet#setHidden} method. The difference is that the
    /// worksheet cannot be unhidden in the the Excel user interface. The Excel
    /// worksheet `xlSheetVeryHidden` option can only be unset programmatically
    /// by VBA.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setVeryHidden", skip_jsdoc)]
    pub fn set_very_hidden(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_very_hidden(enable),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_very_hidden(enable),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set current worksheet as the first visible sheet tab.
    ///
    /// The {@link Worksheet#setActive}  method determines which worksheet is
    /// initially selected. However, if there are a large number of worksheets
    /// the selected worksheet may not appear on the screen. To avoid this you
    /// can select which is the leftmost visible worksheet tab using
    /// `set_first_tab()`.
    ///
    /// This method is not required very often. The default is the first
    /// worksheet.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setFirstTab", skip_jsdoc)]
    pub fn set_first_tab(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_first_tab(enable),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_first_tab(enable),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the color of the worksheet tab.
    ///
    /// The `set_tab_color()` method can be used to change the color of the
    /// worksheet tab. This is useful for highlighting the important tab in a
    /// group of worksheets.
    ///
    /// # Parameters
    ///
    /// - `color`: The tab color property defined by a {@link Color} enum
    ///   value.
    #[wasm_bindgen(js_name = "setTabColor", skip_jsdoc)]
    pub fn set_tab_color(&self, color: Color) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => {
                lock.add_worksheet().set_tab_color(xlsx::Color::from(color))
            }
            WorksheetAccessor::AddChartsheet => lock
                .add_chartsheet()
                .set_tab_color(xlsx::Color::from(color)),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the paper type/size when printing.
    ///
    /// This method is used to set the paper format for the printed output of a
    /// worksheet. The following paper styles are available:
    ///
    /// | Index    | Paper format            | Paper size           |
    /// | :------- | :---------------------- | :------------------- |
    /// | 0        | Printer default         | Printer default      |
    /// | 1        | Letter                  | 8 1/2 x 11 in        |
    /// | 2        | Letter Small            | 8 1/2 x 11 in        |
    /// | 3        | Tabloid                 | 11 x 17 in           |
    /// | 4        | Ledger                  | 17 x 11 in           |
    /// | 5        | Legal                   | 8 1/2 x 14 in        |
    /// | 6        | Statement               | 5 1/2 x 8 1/2 in     |
    /// | 7        | Executive               | 7 1/4 x 10 1/2 in    |
    /// | 8        | A3                      | 297 x 420 mm         |
    /// | 9        | A4                      | 210 x 297 mm         |
    /// | 10       | A4 Small                | 210 x 297 mm         |
    /// | 11       | A5                      | 148 x 210 mm         |
    /// | 12       | B4                      | 250 x 354 mm         |
    /// | 13       | B5                      | 182 x 257 mm         |
    /// | 14       | Folio                   | 8 1/2 x 13 in        |
    /// | 15       | Quarto                  | 215 x 275 mm         |
    /// | 16       | ---                     | 10x14 in             |
    /// | 17       | ---                     | 11x17 in             |
    /// | 18       | Note                    | 8 1/2 x 11 in        |
    /// | 19       | Envelope 9              | 3 7/8 x 8 7/8        |
    /// | 20       | Envelope 10             | 4 1/8 x 9 1/2        |
    /// | 21       | Envelope 11             | 4 1/2 x 10 3/8       |
    /// | 22       | Envelope 12             | 4 3/4 x 11           |
    /// | 23       | Envelope 14             | 5 x 11 1/2           |
    /// | 24       | C size sheet            | ---                  |
    /// | 25       | D size sheet            | ---                  |
    /// | 26       | E size sheet            | ---                  |
    /// | 27       | Envelope DL             | 110 x 220 mm         |
    /// | 28       | Envelope C3             | 324 x 458 mm         |
    /// | 29       | Envelope C4             | 229 x 324 mm         |
    /// | 30       | Envelope C5             | 162 x 229 mm         |
    /// | 31       | Envelope C6             | 114 x 162 mm         |
    /// | 32       | Envelope C65            | 114 x 229 mm         |
    /// | 33       | Envelope B4             | 250 x 353 mm         |
    /// | 34       | Envelope B5             | 176 x 250 mm         |
    /// | 35       | Envelope B6             | 176 x 125 mm         |
    /// | 36       | Envelope                | 110 x 230 mm         |
    /// | 37       | Monarch                 | 3.875 x 7.5 in       |
    /// | 38       | Envelope                | 3 5/8 x 6 1/2 in     |
    /// | 39       | Fanfold                 | 14 7/8 x 11 in       |
    /// | 40       | German Std Fanfold      | 8 1/2 x 12 in        |
    /// | 41       | German Legal Fanfold    | 8 1/2 x 13 in        |
    ///
    /// Note, it is likely that not all of these paper types will be available
    /// to the end user since it will depend on the paper formats that the
    /// user's printer supports. Therefore, it is best to stick to standard
    /// paper types of 1 for US Letter and 9 for A4.
    ///
    /// If you do not specify a paper type the worksheet will print using the
    /// printer's default paper style.
    ///
    /// # Parameters
    ///
    /// - `paper_size`: The paper size index from the list above .
    #[wasm_bindgen(js_name = "setPaperSize", skip_jsdoc)]
    pub fn set_paper_size(&self, paper_size: u8) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_paper_size(paper_size),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_paper_size(paper_size),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the order in which pages are printed.
    ///
    /// The `set_page_order()` method is used to change the default print
    /// direction. This is referred to by Excel as the sheet "page order":
    ///
    /// The default page order is shown below for a worksheet that extends over
    /// 4 pages. The order is called "down then over":
    ///
    /// src="https://rustxlsxwriter.github.io/images/worksheet_set_page_order.png">
    ///
    /// However, by using `set_page_order(false)` the print order will be
    /// changed to "over then down".
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. Set `true` to get "Down, then
    ///   over" (the default) and `false` to get "Over, then down".
    #[wasm_bindgen(js_name = "setPageOrder", skip_jsdoc)]
    pub fn set_page_order(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_page_order(enable),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_page_order(enable),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the page orientation to landscape.
    ///
    /// The `set_landscape()` method is used to set the orientation of a
    /// worksheet's printed page to landscape.
    #[wasm_bindgen(js_name = "setLandscape", skip_jsdoc)]
    pub fn set_landscape(&self) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_landscape(),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_landscape(),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the page orientation to portrait.
    ///
    ///  This `set_portrait()` method  is used to set the orientation of a
    ///  worksheet's printed page to portrait. The default worksheet orientation
    ///  is portrait, so this function is rarely required.
    #[wasm_bindgen(js_name = "setPortrait", skip_jsdoc)]
    pub fn set_portrait(&self) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_portrait(),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_portrait(),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the page view mode to normal layout.
    ///
    /// This method is used to display the worksheet in “View -> Normal”
    /// mode. This is the default.
    #[wasm_bindgen(js_name = "setViewNormal", skip_jsdoc)]
    pub fn set_view_normal(&self) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_view_normal(),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_view_normal(),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the page view mode to page layout.
    ///
    /// This method is used to display the worksheet in “View -> Page Layout”
    /// mode.
    #[wasm_bindgen(js_name = "setViewPageLayout", skip_jsdoc)]
    pub fn set_view_page_layout(&self) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_view_page_layout(),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_view_page_layout(),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the page view mode to page break preview.
    ///
    /// This method is used to display the worksheet in “View -> Page Break
    /// Preview” mode.
    #[wasm_bindgen(js_name = "setViewPageBreakPreview", skip_jsdoc)]
    pub fn set_view_page_break_preview(&self) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_view_page_break_preview(),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_view_page_break_preview(),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the vertical page breaks on a worksheet.
    ///
    /// The `set_vertical_page_breaks()` method adds vertical page breaks to a
    /// worksheet. This is much less common than the
    /// {@link Worksheet#setPageBreaks} method shown above.
    ///
    /// # Parameters
    ///
    /// - `breaks`: A list of one or more column numbers where the page breaks
    ///   occur.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - {@link XlsxError#ParameterError} - The number of page breaks exceeds
    ///   Excel's limit of 1023 page breaks.
    #[wasm_bindgen(js_name = "setVerticalPageBreaks", skip_jsdoc)]
    pub fn set_vertical_page_breaks(&self, breaks: Vec<u32>) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => {
                lock.add_worksheet().set_vertical_page_breaks(&breaks)
            }
            WorksheetAccessor::AddChartsheet => {
                lock.add_chartsheet().set_vertical_page_breaks(&breaks)
            }
        }?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
        })
    }
    /// Set the worksheet zoom factor.
    ///
    /// Set the worksheet zoom factor in the range 10 <= zoom <= 400.
    ///
    /// The default zoom level is 100. The `set_zoom()` method does not affect
    /// the scale of the printed page in Excel. For that you should use
    /// {@link Worksheet#setPrintScale}.
    ///
    /// # Parameters
    ///
    /// - `zoom`: The worksheet zoom level.
    #[wasm_bindgen(js_name = "setZoom", skip_jsdoc)]
    pub fn set_zoom(&self, zoom: u16) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_zoom(zoom),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_zoom(zoom),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set a chartsheet to automatically zoom to fit the screen.
    ///
    /// The `set_zoom_to_fit()` method only has an effect on a
    /// [Chartsheet](Worksheet::new_chartsheet) object. It ensures that the
    /// chartsheet is zoomed automatically by Excel to fit the screen even when
    /// the window is resized.
    ///
    /// This method doesn't have an effect on a standard worksheet.
    ///
    /// # Parameters
    ///
    /// - `enable`: A boolean value to enable or disable the zoom to fit.
    #[wasm_bindgen(js_name = "setZoomToFit", skip_jsdoc)]
    pub fn set_zoom_to_fit(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_zoom_to_fit(enable),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_zoom_to_fit(enable),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the printed page header caption.
    ///
    /// The `set_header()` method can be used to set the header for a worksheet.
    ///
    /// Headers and footers are generated using a string which is a combination
    /// of plain text and optional control characters.
    ///
    /// The available control characters are:
    ///
    /// | Control              | Category      | Description           |
    /// | -------------------- | ------------- | --------------------- |
    /// | `&L`                 | Alignment     | Left                  |
    /// | `&C`                 |               | Center                |
    /// | `&R`                 |               | Right                 |
    /// | `&[Page]`  or `&P`   | Information   | Page number           |
    /// | `&[Pages]` or `&N`   |               | Total number of pages |
    /// | `&[Date]`  or `&D`   |               | Date                  |
    /// | `&[Time]`  or `&T`   |               | Time                  |
    /// | `&[File]`  or `&F`   |               | File name             |
    /// | `&[Tab]`   or `&A`   |               | Worksheet name        |
    /// | `&[Path]`  or `&Z`   |               | Workbook path         |
    /// | `&fontsize`          | Font          | Font size             |
    /// | `&"font,style"`      |               | Font name and style   |
    /// | `&U`                 |               | Single underline      |
    /// | `&E`                 |               | Double underline      |
    /// | `&S`                 |               | Strikethrough         |
    /// | `&X`                 |               | Superscript           |
    /// | `&Y`                 |               | Subscript             |
    /// | `&[Picture]` or `&G` | Images        | Picture/image         |
    /// | `&&`                 | Miscellaneous | Literal ampersand &   |
    ///
    /// Some of the placeholder variables have a long version like `&[Page]` and
    /// a short version like `&P`. The longer version is displayed in the Excel
    /// interface but the shorter version is the way that it is stored in the
    /// file format. Either version is okay since `rust_xlsxwriter` will
    /// translate as required.
    ///
    /// Headers and footers have 3 edit areas to the left, center and right.
    /// Text can be aligned to these areas by prefixing the text with the
    /// control characters `&L`, `&C` and `&R`.
    ///
    /// For example:
    ///
    /// You can also have text in each of the alignment areas:
    ///
    /// The information control characters act as variables/templates that Excel
    /// will update/expand as the workbook or worksheet changes.
    ///
    /// Times and dates are in the user's default format:
    ///
    /// To include a single literal ampersand `&` in a header or footer you
    /// should use a double ampersand `&&`:
    ///
    /// You can specify the font size of a section of the text by prefixing it
    /// with the control character `&n` where `n` is the font size:
    ///
    /// You can specify the font of a section of the text by prefixing it with
    /// the control sequence `&"font,style"` where `fontname` is a font name
    /// such as Windows font descriptions: "Regular", "Italic", "Bold" or "Bold
    /// Italic": "Courier New" or "Times New Roman" and `style` is one of the
    /// standard Windows font descriptions like “Regular”, “Italic”, “Bold” or
    /// “Bold Italic”:
    ///
    /// It is possible to combine all of these features together to create
    /// complex headers and footers. If you set up a complex header in Excel you
    /// can transfer it to `rust_xlsxwriter` by inspecting the string in the
    /// Excel file. For example the following shows how unzip and grep the Excel
    /// XML sub-files on a Linux system. The example uses libxml's xmllint to
    /// format the XML for clarity:
    ///
    /// Note: Excel requires that the header or footer string be less than 256
    /// characters, including the control characters. Strings longer than this
    /// will not be written, and a warning will be output.
    ///
    /// # Parameters
    ///
    /// - `header`: The header string with optional control characters.
    #[wasm_bindgen(js_name = "setHeader", skip_jsdoc)]
    pub fn set_header(&self, header: &str) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_header(header),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_header(header),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the printed page footer caption.
    ///
    /// The `set_footer()` method can be used to set the footer for a worksheet.
    ///
    /// See the documentation for {@link Worksheet#setHeader} for more details
    /// on the syntax of the header/footer string.
    ///
    /// # Parameters
    ///
    /// - `footer`: The footer string with optional control characters.
    #[wasm_bindgen(js_name = "setFooter", skip_jsdoc)]
    pub fn set_footer(&self, footer: &str) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_footer(footer),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_footer(footer),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Insert an image in a worksheet header.
    ///
    /// Insert an image in a worksheet header in one of the 3 sections supported
    /// by Excel: Left, Center and Right. This needs to be preceded by a call to
    /// {@link Worksheet#setHeader} where a corresponding `&[Picture]` element
    /// is added to the header formatting string such as `"&L&[Picture]"`.
    ///
    /// # Parameters
    ///
    /// - `position`: The image position as defined by the
    ///   {@link HeaderImagePosition} enum.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#ParameterError} - Parameter error if there isn't a
    ///   corresponding `&[Picture]`/`&[G]` variable in the header string.
    #[wasm_bindgen(js_name = "setHeaderImage", skip_jsdoc)]
    pub fn set_header_image(
        &self,
        image: Image,
        position: HeaderImagePosition,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_header_image(
                &*image.inner.lock().unwrap(),
                xlsx::HeaderImagePosition::from(position),
            ),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_header_image(
                &*image.inner.lock().unwrap(),
                xlsx::HeaderImagePosition::from(position),
            ),
        }?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
        })
    }
    /// Insert an image in a worksheet footer.
    ///
    /// See the documentation for {@link Worksheet#setHeaderImage} for more
    /// details.
    ///
    /// # Parameters
    ///
    /// - `position`: The image position as defined by the
    ///   {@link HeaderImagePosition} enum.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#ParameterError} - Parameter error if there isn't a
    ///   corresponding `&[Picture]`/`&[G]` variable in the header string.
    #[wasm_bindgen(js_name = "setFooterImage", skip_jsdoc)]
    pub fn set_footer_image(
        &self,
        image: Image,
        position: HeaderImagePosition,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_footer_image(
                &*image.inner.lock().unwrap(),
                xlsx::HeaderImagePosition::from(position),
            ),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_footer_image(
                &*image.inner.lock().unwrap(),
                xlsx::HeaderImagePosition::from(position),
            ),
        }?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
        })
    }
    /// Set the page setup option to scale the header/footer with the document.
    ///
    /// This option determines whether the headers and footers use the same
    /// scaling as the worksheet. This defaults to "on" in Excel.
    ///
    /// See also the documentation on [Worksheet Page Setup -
    /// Header/Footer](../worksheet/index.html#page-setup---headerfooter).
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is on by default.
    #[wasm_bindgen(js_name = "setHeaderFooterScaleWithDoc", skip_jsdoc)]
    pub fn set_header_footer_scale_with_doc(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock
                .add_worksheet()
                .set_header_footer_scale_with_doc(enable),
            WorksheetAccessor::AddChartsheet => lock
                .add_chartsheet()
                .set_header_footer_scale_with_doc(enable),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the page setup option to align the header/footer with the margins.
    ///
    /// This option determines whether the headers and footers align with the
    /// left and right margins of the worksheet. This defaults to "on" in Excel.
    ///
    /// See also the documentation on [Worksheet Page Setup -
    /// Header/Footer](../worksheet/index.html#page-setup---headerfooter).
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is on by default. S
    #[wasm_bindgen(js_name = "setHeaderFooterAlignWithPage", skip_jsdoc)]
    pub fn set_header_footer_align_with_page(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock
                .add_worksheet()
                .set_header_footer_align_with_page(enable),
            WorksheetAccessor::AddChartsheet => lock
                .add_chartsheet()
                .set_header_footer_align_with_page(enable),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the page margins.
    ///
    /// The `set_margins()` method is used to set the margins of the worksheet
    /// when it is printed. The units are in inches. Specifying `-1.0` for any
    /// parameter will give the default Excel value. The defaults are shown
    /// below.
    ///
    /// # Parameters
    ///
    /// - `left`: Left margin in inches. Excel default is 0.7.
    /// - `right`: Right margin in inches. Excel default is 0.7.
    /// - `top`: Top margin in inches. Excel default is 0.75.
    /// - `bottom`: Bottom margin in inches. Excel default is 0.75.
    /// - `header`: Header margin in inches. Excel default is 0.3.
    /// - `footer`: Footer margin in inches. Excel default is 0.3.
    #[wasm_bindgen(js_name = "setMargins", skip_jsdoc)]
    pub fn set_margins(
        &self,
        left: f64,
        right: f64,
        top: f64,
        bottom: f64,
        header: f64,
        footer: f64,
    ) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock
                .add_worksheet()
                .set_margins(left, right, top, bottom, header, footer),
            WorksheetAccessor::AddChartsheet => lock
                .add_chartsheet()
                .set_margins(left, right, top, bottom, header, footer),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the first page number when printing.
    ///
    /// The `set_print_first_page_number()` method is used to set the page
    /// number of the first page when the worksheet is printed out. This option
    /// will only have and effect if you have a header/footer with the `&[Page]`
    /// control character, see {@link Worksheet#setHeader}.
    ///
    /// See also the documentation on [Worksheet Page Setup -
    /// Page](../worksheet/index.html#page-setup---page).
    ///
    /// # Parameters
    ///
    /// - `page_number`: The page number of the first printed page.
    #[wasm_bindgen(js_name = "setPrintFirstPageNumber", skip_jsdoc)]
    pub fn set_print_first_page_number(&self, page_number: u16) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock
                .add_worksheet()
                .set_print_first_page_number(page_number),
            WorksheetAccessor::AddChartsheet => lock
                .add_chartsheet()
                .set_print_first_page_number(page_number),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the page setup option to set the print scale.
    ///
    /// Set the scale factor of the printed page, in the range 10 <= scale <=
    /// 400.
    ///
    /// The default scale factor is 100. The `set_print_scale()` method
    /// does not affect the scale of the visible page in Excel. For that you
    /// should use {@link Worksheet#setZoom}.
    ///
    /// See also the documentation on [Worksheet Page Setup -
    /// Page](../worksheet/index.html#page-setup---page).
    ///
    /// # Parameters
    ///
    /// - `scale`: The print scale factor.
    #[wasm_bindgen(js_name = "setPrintScale", skip_jsdoc)]
    pub fn set_print_scale(&self, scale: u16) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_print_scale(scale),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_print_scale(scale),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Fit the printed area to a specific number of pages both vertically and
    /// horizontally.
    ///
    /// The `set_print_fit_to_pages()` method is used to fit the printed area to
    /// a specific number of pages both vertically and horizontally. If the
    /// printed area exceeds the specified number of pages it will be scaled
    /// down to fit. This ensures that the printed area will always appear on
    /// the specified number of pages even if the page size or margins change:
    ///
    /// The print area can be defined using the `set_print_area()` method.
    ///
    /// A common requirement is to fit the printed output to `n` pages wide but
    /// have the height be as long as necessary. To achieve this set the
    /// `height` to 0, see the example below.
    ///
    /// **Notes**:
    ///
    /// - The `set_print_fit_to_pages()` will override any manual page breaks
    ///   that are defined in the worksheet.
    ///
    /// - When using `set_print_fit_to_pages()` it may also be required to set
    ///   the printer paper size using {@link Worksheet#setPaperSize} or else
    ///   Excel will default to "US Letter".
    ///
    /// See also the documentation on [Worksheet Page Setup -
    /// Page](../worksheet/index.html#page-setup---page).
    ///
    /// # Parameters
    ///
    /// - `width`: Number of pages horizontally.
    /// - `height`: Number of pages vertically.
    #[wasm_bindgen(js_name = "setPrintFitToPages", skip_jsdoc)]
    pub fn set_print_fit_to_pages(&self, width: u16, height: u16) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => {
                lock.add_worksheet().set_print_fit_to_pages(width, height)
            }
            WorksheetAccessor::AddChartsheet => {
                lock.add_chartsheet().set_print_fit_to_pages(width, height)
            }
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Center the printed page horizontally.
    ///
    /// Center the worksheet data horizontally between the margins on the
    /// printed page
    ///
    /// See also the documentation on [Worksheet Page Setup -
    /// Margins](../worksheet/index.html#page-setup---margins).
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setPrintCenterHorizontally", skip_jsdoc)]
    pub fn set_print_center_horizontally(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => {
                lock.add_worksheet().set_print_center_horizontally(enable)
            }
            WorksheetAccessor::AddChartsheet => {
                lock.add_chartsheet().set_print_center_horizontally(enable)
            }
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Center the printed page vertically.
    ///
    /// Center the worksheet data vertically between the margins on the printed
    /// page
    ///
    /// See also the documentation on [Worksheet Page Setup -
    /// Margins](../worksheet/index.html#page-setup---margins).
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setPrintCenterVertically", skip_jsdoc)]
    pub fn set_print_center_vertically(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => {
                lock.add_worksheet().set_print_center_vertically(enable)
            }
            WorksheetAccessor::AddChartsheet => {
                lock.add_chartsheet().set_print_center_vertically(enable)
            }
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the option to turn on/off the screen gridlines.
    ///
    /// The `set_screen_gridlines()` method is use to turn on/off gridlines on
    /// displayed worksheet. It is on by default.
    ///
    /// To turn on/off the printed gridlines see the
    /// {@link Worksheet#setPrintGridlines} method below.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is on by default.
    #[wasm_bindgen(js_name = "setScreenGridlines", skip_jsdoc)]
    pub fn set_screen_gridlines(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_screen_gridlines(enable),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_screen_gridlines(enable),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the page setup option to turn on printed gridlines.
    ///
    /// The `set_print_gridlines()` method is use to turn on/off gridlines on
    /// the printed pages. It is off by default.
    ///
    /// To turn on/off the screen gridlines see the
    /// {@link Worksheet#setScreenGridlines} method above.
    ///
    /// See also the documentation on [Worksheet Page Setup -
    /// Sheet](../worksheet/index.html#page-setup---sheet).
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setPrintGridlines", skip_jsdoc)]
    pub fn set_print_gridlines(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_print_gridlines(enable),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_print_gridlines(enable),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the page setup option to print in black and white.
    ///
    /// This `set_print_black_and_white()` method can be used to force printing
    /// in black and white only. It is off by default.
    ///
    /// See also the documentation on [Worksheet Page Setup -
    /// Sheet](../worksheet/index.html#page-setup---sheet).
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setPrintBlackAndWhite", skip_jsdoc)]
    pub fn set_print_black_and_white(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => {
                lock.add_worksheet().set_print_black_and_white(enable)
            }
            WorksheetAccessor::AddChartsheet => {
                lock.add_chartsheet().set_print_black_and_white(enable)
            }
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the page setup option to print in draft quality.
    ///
    /// See also the documentation on [Worksheet Page Setup -
    /// Sheet](../worksheet/index.html#page-setup---sheet).
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setPrintDraft", skip_jsdoc)]
    pub fn set_print_draft(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_print_draft(enable),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_print_draft(enable),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the page setup option to print the row and column headers on the
    /// printed page.
    ///
    /// The `set_print_headings()` method turns on the row and column headers
    /// when printing a worksheet. This option is off by default.
    ///
    /// See also the documentation on [Worksheet Page Setup -
    /// Sheet](../worksheet/index.html#page-setup---sheet).
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setPrintHeadings", skip_jsdoc)]
    pub fn set_print_headings(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_print_headings(enable),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_print_headings(enable),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Autofit the worksheet column widths to the widest data in the column,
    /// approximately.
    ///
    /// Excel autofits columns at runtime when it has access to all of the
    /// required worksheet information as well as the Windows functions for
    /// calculating display areas based on fonts and formatting.
    ///
    /// The `rust_xlsxwriter` library doesn't have access to these Windows
    /// functions so it simulates autofit by calculating string widths based on
    /// font character metrics taken from Excel.
    ///
    /// This isn't perfect but for most cases it should be sufficient and
    /// indistinguishable from the output of Excel. However there are some
    /// limitations to be aware of when using this method:
    ///
    /// - It is based on the default Excel font type and size of Calibri 11. It
    ///   will not give accurate results for other fonts or font sizes.
    /// - It only takes formatting of numbers or dates into account if the
    ///   `enhanced_autofit` feature is enabled, which requires the {@link ssfmt}
    ///   crate. See the second example below.
    /// - Autofit is a relatively expensive operation since it performs a
    ///   calculation for all the populated cells in a worksheet. This can be
    ///   mitigated by using the {@link Worksheet#setAutofitMaxRow} method,
    ///   see below.
    ///
    /// For cases that don't match your desired output you can set explicit
    /// column widths via {@link Worksheet#setColumnWidth} or
    /// {@link Worksheet#setColumnWidthPixels}. The `autofit()` method ignores
    /// columns that have already been explicitly set if the width is greater
    /// than the calculated autofit width. Alternatively, setting the column
    /// width explicitly after calling `autofit()` will override the autofit
    /// value.
    #[wasm_bindgen(js_name = "autofit", skip_jsdoc)]
    pub fn autofit(&self) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().autofit(),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().autofit(),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the maximum autofit width for worksheet columns.
    ///
    /// The {@link Worksheet#autofit} method simulates Excel's column autofit.
    /// One undesirable side-effect of this is that Excel autofits very long
    /// strings up to limit of 255 characters/1790 pixels. This is often too
    /// wide to display on a single screen at normal zoom. As such the
    /// `set_autofit_max_width()` method can be used to set a smaller upper
    /// limit for autofitting long strings.
    ///
    /// A value of 300 pixels is recommended as a good compromise between column
    /// width and readability.
    ///
    /// # Parameters
    ///
    /// - `max_width`: The maximum column width, in pixels, to use for autofitting.
    #[wasm_bindgen(js_name = "setAutofitMaxWidth", skip_jsdoc)]
    pub fn set_autofit_max_width(&self, max_width: u32) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => {
                lock.add_worksheet().set_autofit_max_width(max_width)
            }
            WorksheetAccessor::AddChartsheet => {
                lock.add_chartsheet().set_autofit_max_width(max_width)
            }
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Autofit the worksheet columns up to a maximum width.
    ///
    /// The {@link Worksheet#autofit} method above simulates Excel's column
    /// autofit. One undesirable side-effect of this is that Excel autofits very
    /// long strings up to limit of 255 characters/1790 pixels. This is often
    /// too wide to display on a single screen at normal zoom. As such the
    /// `autofit_to_max_width()` method is provided to enable a smaller upper
    /// limit for autofitting long strings. A value of 300 pixels is recommended
    /// as a good compromise between column width and readability.
    ///
    /// # Parameters
    ///
    /// - `max_width`: The maximum column width, in pixels, to use for
    ///   autofitting.
    #[wasm_bindgen(js_name = "autofitToMaxWidth", skip_jsdoc)]
    pub fn autofit_to_max_width(&self, max_width: u32) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().autofit_to_max_width(max_width),
            WorksheetAccessor::AddChartsheet => {
                lock.add_chartsheet().autofit_to_max_width(max_width)
            }
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the worksheet name used in VBA macros.
    ///
    /// This method can be used to set the VBA name for the worksheet. This is
    /// sometimes required when a VBA macro included via
    /// {@link Workbook#addVbaProject})
    /// makes reference to the worksheet with a name other than the default
    /// Excel VBA names of `Sheet1`, `Sheet2`, etc.
    ///
    /// See also the
    /// {@link Workbook#setVbaName}) method for
    /// setting the workbook VBA name.
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
    /// The name must be also be unique across the worksheets in the workbook.
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
    pub fn set_vba_name(&self, name: &str) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_vba_name(name),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_vba_name(name),
        }?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
        })
    }
    /// Set the default string used for NaN values.
    ///
    /// Excel doesn't support storing `NaN` (Not a Number) values. If a `NAN` is
    /// generated as the result of a calculation Excel stores and displays the
    /// error `#NUM!`. However, this error isn't usually used outside of a
    /// formula result and it isn't stored as a number.
    ///
    /// In order to deal with `f64::NAN` in a reasonable way `rust_xlsxwriter`
    /// writes it as the string "NAN". The `set_nan_value()` method allows you
    /// to override this default value.
    ///
    /// # Parameters
    ///
    /// - `value`: The string to use for NaN values.
    #[wasm_bindgen(js_name = "setNanValue", skip_jsdoc)]
    pub fn set_nan_value(&self, value: &str) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_nan_value(value),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_nan_value(value),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the default string used for Infinite values.
    ///
    /// Excel doesn't support storing `Inf` (infinity) values. If an `Inf` is
    /// generated as the result of a calculation Excel stores and displays the
    /// error `#DIV/0`. However, this error isn't usually used outside of a
    /// formula result and it isn't stored as a number.
    ///
    /// In order to deal with `f64::INFINITY` in a reasonable way
    /// `rust_xlsxwriter` writes it as the string "INF". The
    /// `set_infinite_value()` method allows you to override this default value.
    ///
    /// See the example for {@link Worksheet#setNanValue} above.
    ///
    /// # Parameters
    ///
    /// - `value`: The string to use for `Inf` values.
    #[wasm_bindgen(js_name = "setInfinityValue", skip_jsdoc)]
    pub fn set_infinity_value(&self, value: &str) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_infinity_value(value),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_infinity_value(value),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the default string used for negative Infinite values.
    ///
    /// Excel doesn't support storing `-Inf` (negative infinity) values. If a
    /// `-Inf` is generated as the result of a calculation Excel stores and
    /// displays the error `#DIV/0`. However, this error isn't usually used
    /// outside of a formula result and it isn't stored as a number.
    ///
    /// In order to deal with `f64::NEG_INFINITY` in a reasonable way
    /// `rust_xlsxwriter` writes it as the string "-INF". The
    /// `set_infinite_value()` method method allows you to override this default
    /// value.
    ///
    /// See the example for {@link Worksheet#setNanValue} above.
    ///
    /// # Parameters
    ///
    /// - `value`: The string to use for `-Inf` values.
    #[wasm_bindgen(js_name = "setNegInfinityValue", skip_jsdoc)]
    pub fn set_neg_infinity_value(&self, value: &str) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_neg_infinity_value(value),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_neg_infinity_value(value),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Get the local instance DXF id for a format.
    ///
    /// Get the local instance DXF id for a format. These indexes will be
    /// replaced by global/workbook indices before the worksheet is saved. DXF
    /// indexed are used for Tables and Conditional Formats.
    ///
    /// This method is public but hidden to allow test cases to mirror the
    /// creation order for DXF ids which is usually the reverse of the order of
    /// the XF instance ids.
    ///
    /// # Parameters
    ///
    /// `format` - The {@link Format} instance to register.
    #[wasm_bindgen(js_name = "formatDxfIndex", skip_jsdoc)]
    pub fn format_dxf_index(&self, format: Format) -> u32 {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock
                .add_worksheet()
                .format_dxf_index(&*format.inner.lock().unwrap()),
            WorksheetAccessor::AddChartsheet => lock
                .add_chartsheet()
                .format_dxf_index(&*format.inner.lock().unwrap()),
        }
    }
}
