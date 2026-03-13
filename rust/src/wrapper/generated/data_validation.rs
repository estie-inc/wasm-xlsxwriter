use crate::wrapper::DataValidationErrorStyle;
use crate::wrapper::Formula;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `DataValidation` struct represents a data validation in Excel.
///
/// `DataValidation` is used in conjunction with the
/// {@link Worksheet#addDataValidation}
/// method.
#[derive(Clone)]
#[wasm_bindgen]
pub struct DataValidation {
    pub(crate) inner: Arc<Mutex<xlsx::DataValidation>>,
}

#[wasm_bindgen]
impl DataValidation {
    #[wasm_bindgen(constructor)]
    pub fn new() -> DataValidation {
        DataValidation {
            inner: Arc::new(Mutex::new(xlsx::DataValidation::new())),
        }
    }
    #[doc = r" Create a deep clone of this object."]
    #[wasm_bindgen(js_name = "clone")]
    pub fn deep_clone(&self) -> DataValidation {
        DataValidation {
            inner: Arc::new(Mutex::new(self.inner.lock().unwrap().clone())),
        }
    }
    /// Set a data validation rule to restrict cell input to a selection of
    /// strings via a dropdown menu and a cell range reference.
    ///
    /// The strings for the dropdown should be placed somewhere in the
    /// worksheet/workbook and should be referred to via a cell range reference
    /// represented by a {@link Formula}, see [Using cell references in Data
    /// Validations] and the example below.
    ///
    /// # Parameters
    /// - `list`: A cell range reference such as `=B1:B9`, `=$B$1:$B$9` or
    ///   `=Sheet2!B1:B9` using a {@link Formula}. See [Using cell references in
    ///   Data Validations].
    ///
    /// [Using cell references in Data Validations]:
    ///     #using-cell-references-in-data-validations
    #[wasm_bindgen(js_name = "allowListFormula", skip_jsdoc)]
    pub fn allow_list_formula(&self, list: Formula) -> DataValidation {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::DataValidation::new());
        inner = inner.allow_list_formula(list.inner.lock().unwrap().clone());
        *lock = inner;
        DataValidation {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set a data validation to limit input based on a custom formula.
    ///
    /// Set a data validation rule to restrict cell input based on an Excel
    /// formula that returns a boolean value. Excel refers to this data
    /// validation type as "Custom".
    ///
    /// # Parameters
    ///
    /// - `rule`: A {@link Formula} value. You should ensure that the formula is
    ///   valid in Excel.
    #[wasm_bindgen(js_name = "allowCustom", skip_jsdoc)]
    pub fn allow_custom(&self, rule: Formula) -> DataValidation {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::DataValidation::new());
        inner = inner.allow_custom(rule.inner.lock().unwrap().clone());
        *lock = inner;
        DataValidation {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set a data validation to allow any input data.
    ///
    /// The "Any" data validation type doesn't restrict data input and is mainly
    /// used to allow access to the "Input Message" dialog when a user enters data
    /// in a cell.
    ///
    /// This is the default validation type for {@link DataValidation} if no other
    /// `allow_TYPE()` method is used. Situations where this type of data
    /// validation are required are uncommon.
    #[wasm_bindgen(js_name = "allowAnyValue", skip_jsdoc)]
    pub fn allow_any_value(&self) -> DataValidation {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::DataValidation::new());
        inner = inner.allow_any_value();
        *lock = inner;
        DataValidation {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the data validation option that defines how blank cells are handled.
    ///
    /// By default, Excel data validations have an "Ignore blank" option turned
    /// on. This allows the user to optionally leave the cell blank and not
    /// enter any value. This is generally the best default option since it
    /// allows users to exit the cell without inputting any data.
    ///
    /// src="https://rustxlsxwriter.github.io/images/data_validation_allow_whole_number_dialog.png">
    ///
    /// If you need to ensure that the user inserts some information, then you
    /// can use `ignore_blank()` to turn this option off.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is on by default.
    #[wasm_bindgen(js_name = "ignoreBlank", skip_jsdoc)]
    pub fn ignore_blank(&self, enable: bool) -> DataValidation {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::DataValidation::new());
        inner = inner.ignore_blank(enable);
        *lock = inner;
        DataValidation {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Turn on/off the in-cell dropdown for list data validations.
    ///
    /// By default, the Excel list data validation has an "In-cell drop-down"
    /// option turned on. This shows a dropdown arrow for list-style data
    /// validations and displays the list items.
    ///
    /// src="https://rustxlsxwriter.github.io/images/data_validation_allow_list_strings_dialog.png">
    ///
    /// If this option is turned off the data validation will restrict input to
    /// the specified list values but it won't display a visual indicator of
    /// what those values are.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is on by default.
    #[wasm_bindgen(js_name = "showDropdown", skip_jsdoc)]
    pub fn show_dropdown(&self, enable: bool) -> DataValidation {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::DataValidation::new());
        inner = inner.show_dropdown(enable);
        *lock = inner;
        DataValidation {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Toggle option to show an input message when the cell is entered.
    ///
    /// This function is used to toggle the option that controls whether an
    /// input message is shown when a data validation cell is entered.
    ///
    /// The option only has an effect if there is an input message, so for the
    /// majority of use cases it isn't required.
    ///
    /// See also {@link DataValidation#setInputMessage} below.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is on by default.
    #[wasm_bindgen(js_name = "showInputMessage", skip_jsdoc)]
    pub fn show_input_message(&self, enable: bool) -> DataValidation {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::DataValidation::new());
        inner = inner.show_input_message(enable);
        *lock = inner;
        DataValidation {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the title for the input message when the cell is entered.
    ///
    /// This option is used to set a title in bold for the input message when a
    /// data validation cell is entered.
    ///
    /// The title is only visible if there is also an input message. See the
    /// {@link DataValidation#setInputMessage} example below.
    ///
    /// # Parameters
    ///
    /// - `text`: Title string. Must be less than or equal to the Excel limit
    ///   of 32 characters.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#DataValidationError} - The length of the title exceeds
    ///   Excel's limit of 32 characters.
    #[wasm_bindgen(js_name = "setInputTitle", skip_jsdoc)]
    pub fn set_input_title(&self, text: &str) -> WasmResult<DataValidation> {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::DataValidation::new());
        inner = inner.set_input_title(text)?;
        *lock = inner;
        Ok(DataValidation {
            inner: Arc::clone(&self.inner),
        })
    }
    /// Set the input message when a data validation cell is entered.
    ///
    /// This option is used to set an input message when a data validation cell
    /// is entered. This can be used to explain to the user what the data
    /// validation rules are for the cell.
    ///
    /// # Parameters
    ///
    /// - `text`: Message string. Must be less than or equal to the Excel limit
    ///   of 255 characters. The string can contain newlines to split it over
    ///   several lines.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#DataValidationError} - The length of the message exceeds
    ///   Excel's limit of 255 characters.
    #[wasm_bindgen(js_name = "setInputMessage", skip_jsdoc)]
    pub fn set_input_message(&self, text: &str) -> WasmResult<DataValidation> {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::DataValidation::new());
        inner = inner.set_input_message(text)?;
        *lock = inner;
        Ok(DataValidation {
            inner: Arc::clone(&self.inner),
        })
    }
    /// Toggle option to show an error message when there is a validation error.
    ///
    /// This function is used to toggle the option that controls whether an
    /// error message is shown when there is a validation error.
    ///
    /// If this option is toggled off then any data can be entered in a cell and
    /// an error message will not be raised, which has limited practical
    /// applications for a data validation.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is on by default.
    #[wasm_bindgen(js_name = "showErrorMessage", skip_jsdoc)]
    pub fn show_error_message(&self, enable: bool) -> DataValidation {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::DataValidation::new());
        inner = inner.show_error_message(enable);
        *lock = inner;
        DataValidation {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the title for the error message when there is a validation error.
    ///
    /// This option is used to set a title in bold for the error message when
    /// there is a validation error.
    ///
    /// # Parameters
    ///
    /// - `text`: Title string. Must be less than or equal to the Excel limit
    ///   of 32 characters.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#DataValidationError} - The length of the title exceeds
    ///   Excel's limit of 32 characters.
    #[wasm_bindgen(js_name = "setErrorTitle", skip_jsdoc)]
    pub fn set_error_title(&self, text: &str) -> WasmResult<DataValidation> {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::DataValidation::new());
        inner = inner.set_error_title(text)?;
        *lock = inner;
        Ok(DataValidation {
            inner: Arc::clone(&self.inner),
        })
    }
    /// Set the error message when there is a validation error.
    ///
    /// This option is used to set an error message when there is a validation
    /// error. This can be used to explain to the user what the data validation
    /// rules are for the cell.
    ///
    /// # Parameters
    ///
    /// - `text`: Message string. Must be less than or equal to the Excel limit
    ///   of 255 characters. The string can contain newlines to split it over
    ///   several lines.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#DataValidationError} - The length of the message exceeds
    ///   Excel's limit of 255 characters.
    #[wasm_bindgen(js_name = "setErrorMessage", skip_jsdoc)]
    pub fn set_error_message(&self, text: &str) -> WasmResult<DataValidation> {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::DataValidation::new());
        inner = inner.set_error_message(text)?;
        *lock = inner;
        Ok(DataValidation {
            inner: Arc::clone(&self.inner),
        })
    }
    /// Set the style of the error dialog type for a data validation.
    ///
    /// Set the error dialog to be either "Stop" (the default), "Warning" or
    /// "Information". This option only has an effect on Windows.
    ///
    /// # Parameters
    ///
    /// - `error_style`: A {@link DataValidationErrorStyle} enum value.
    #[wasm_bindgen(js_name = "setErrorStyle", skip_jsdoc)]
    pub fn set_error_style(&self, error_style: DataValidationErrorStyle) -> DataValidation {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::DataValidation::new());
        inner = inner.set_error_style(xlsx::DataValidationErrorStyle::from(error_style));
        *lock = inner;
        DataValidation {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set an additional multi-cell range for the data validation.
    ///
    /// The `set_multi_range()` method is used to extend a data validation
    /// over non-contiguous ranges like `"B3 I3 B9:D12 I9:K12"`.
    ///
    /// # Parameters
    ///
    /// - `range`: A string-like type representing an Excel range.
    #[wasm_bindgen(js_name = "setMultiRange", skip_jsdoc)]
    pub fn set_multi_range(&self, range: &str) -> DataValidation {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::DataValidation::new());
        inner = inner.set_multi_range(range);
        *lock = inner;
        DataValidation {
            inner: Arc::clone(&self.inner),
        }
    }
}
