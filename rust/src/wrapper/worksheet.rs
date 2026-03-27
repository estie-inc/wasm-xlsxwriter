use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

use crate::error::XlsxError;
use crate::wrapper::{
    datetime::ExcelDateTime,
    excel_data::ExcelData,
    excel_data::{JsExcelData, JsExcelDataArray, JsExcelDataMatrix},
    rich_string::RichString,
    utils, ConditionalFormat2ColorScale, ConditionalFormat3ColorScale,
    ConditionalFormatAverage, ConditionalFormatBlank, ConditionalFormatCell,
    ConditionalFormatDataBar, ConditionalFormatDate, ConditionalFormatDuplicate,
    ConditionalFormatError, ConditionalFormatFormula, ConditionalFormatIconSet,
    ConditionalFormatText, ConditionalFormatTop, Format, Formula, Url, WasmResult, Worksheet,
};

/// Companion file: custom methods that cannot be auto-generated.
/// The struct definition and auto-generated methods are in generated/worksheet.rs.
#[wasm_bindgen]
impl Worksheet {
    /// Write generic data to a cell.
    ///
    /// The `write()` method writes data of type {@link ExcelData} to a worksheet.
    ///
    /// @param {number} row - The zero indexed row number.
    /// @param {number} col - The zero indexed column number.
    /// @param {ExcelData} data - Data to write.
    /// @returns {Worksheet} - The worksheet object.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::MaxStringLengthExceeded`] - String exceeds Excel's limit
    ///   of 32,767 characters.
    #[wasm_bindgen(js_name = "write", skip_jsdoc)]
    pub fn write(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        data: &JsExcelData,
    ) -> WasmResult<Worksheet> {
        let mut book = self.parent.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let data: ExcelData = data.try_into()?;
        let _ = sheet.write(row, col, data)?;
        Ok(self.clone())
    }

    /// Write formatted generic data to a cell.
    ///
    /// The `writeWithFormat()` method writes data of type {@link ExcelData} to a worksheet.
    ///
    /// See {@link Worksheet#write} for a list of supported data types.
    /// See {@link Format} for a list of supported formatting options.
    ///
    /// @param {number} row - The zero indexed row number.
    /// @param {number} col - The zero indexed column number.
    /// @param {ExcelData} data - Data to write.
    /// @param {Format} format - The {@link Format} property for the cell.
    /// @returns {Worksheet} - The worksheet object.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::MaxStringLengthExceeded`] - String exceeds Excel's limit
    ///   of 32,767 characters.
    #[wasm_bindgen(js_name = "writeWithFormat", skip_jsdoc)]
    pub fn write_with_format(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        data: &JsExcelData,
        format: &Format,
    ) -> WasmResult<Worksheet> {
        let mut book = self.parent.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let data: ExcelData = data.try_into()?;
        let _ = sheet.write_with_format(row, col, data, &format.inner.lock().unwrap())?;
        Ok(self.clone())
    }

    #[wasm_bindgen(js_name = "writeRichString")]
    pub fn write_rich_string(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        rich_string: &RichString,
    ) -> WasmResult<Worksheet> {
        let mut book = self.parent.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let rich_string = rich_string.lock();
        let rich_string: Vec<_> = rich_string.iter().map(|(f, s)| (f, s.as_str())).collect();
        let _ = sheet.write_rich_string(row, col, &rich_string)?;
        Ok(self.clone())
    }

    #[wasm_bindgen(js_name = "writeRichStringWithFormat")]
    pub fn write_rich_string_with_format(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        rich_string: &RichString,
        format: &Format,
    ) -> WasmResult<Worksheet> {
        let mut book = self.parent.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let rich_string = rich_string.lock();
        let rich_string: Vec<_> = rich_string.iter().map(|(f, s)| (f, s.as_str())).collect();
        let _ = sheet.write_rich_string_with_format(
            row,
            col,
            &rich_string,
            &format.inner.lock().unwrap(),
        )?;
        Ok(self.clone())
    }

    #[wasm_bindgen(js_name = "writeColumn")]
    pub fn write_column(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        values: &JsExcelDataArray,
    ) -> WasmResult<Worksheet> {
        let mut book = self.parent.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let values: Vec<ExcelData> = values.try_into()?;
        let _ = sheet.write_column(row, col, values)?;
        Ok(self.clone())
    }

    #[wasm_bindgen(js_name = "writeColumnWithFormat")]
    pub fn write_column_with_format(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        values: &JsExcelDataArray,
        format: &Format,
    ) -> WasmResult<Worksheet> {
        let mut book = self.parent.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let values: Vec<ExcelData> = values.try_into()?;
        let _ = sheet.write_column_with_format(
            row,
            col,
            values,
            &format.inner.lock().unwrap(),
        )?;
        Ok(self.clone())
    }

    #[wasm_bindgen(js_name = "writeColumnMatrix")]
    pub fn write_column_matrix(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        data: &JsExcelDataMatrix,
    ) -> WasmResult<Worksheet> {
        let mut book = self.parent.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let values: Vec<Vec<ExcelData>> = data.try_into()?;
        let _ = sheet.write_column_matrix(row, col, values)?;
        Ok(self.clone())
    }

    #[wasm_bindgen(js_name = "writeRow")]
    pub fn write_row(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        values: &JsExcelDataArray,
    ) -> WasmResult<Worksheet> {
        let mut book = self.parent.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let values: Vec<ExcelData> = values.try_into()?;
        let _ = sheet.write_row(row, col, values)?;
        Ok(self.clone())
    }

    #[wasm_bindgen(js_name = "writeRowWithFormat")]
    pub fn write_row_with_format(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        values: &JsExcelDataArray,
        format: &Format,
    ) -> WasmResult<Worksheet> {
        let mut book = self.parent.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let values: Vec<ExcelData> = values.try_into()?;
        let _ = sheet.write_row_with_format(
            row,
            col,
            values,
            &format.inner.lock().unwrap(),
        )?;
        Ok(self.clone())
    }

    #[wasm_bindgen(js_name = "writeRowMatrix")]
    pub fn write_row_matrix(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        data: &JsExcelDataMatrix,
    ) -> WasmResult<Worksheet> {
        let mut book = self.parent.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let values: Vec<Vec<ExcelData>> = data.try_into()?;
        let _ = sheet.write_row_matrix(row, col, values)?;
        Ok(self.clone())
    }

    /// Accepts both JS `Date` and `ExcelDateTime`.
    #[wasm_bindgen(js_name = "writeDatetime", skip_jsdoc)]
    pub fn write_datetime(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        datetime: &JsValue,
    ) -> WasmResult<Worksheet> {
        let mut book = self.parent.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        if let Some(dt) = utils::datetime_of_jsval(datetime.clone()) {
            let _ = sheet.write_datetime(row, col, dt)?;
            Ok(self.clone())
        } else if let Some(dt) = utils::excel_datetime_of_jsval(datetime) {
            let _ = sheet.write_datetime(row, col, dt)?;
            Ok(self.clone())
        } else {
            Err(XlsxError::InvalidDate)
        }
    }

    /// Accepts both JS `Date` and `ExcelDateTime`.
    #[wasm_bindgen(js_name = "writeDatetimeWithFormat", skip_jsdoc)]
    pub fn write_datetime_with_format(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        datetime: &JsValue,
        format: &Format,
    ) -> WasmResult<Worksheet> {
        let mut book = self.parent.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        if let Some(dt) = utils::datetime_of_jsval(datetime.clone()) {
            let _ = sheet.write_datetime_with_format(
                row,
                col,
                dt,
                &format.inner.lock().unwrap(),
            )?;
            Ok(self.clone())
        } else if let Some(dt) = utils::excel_datetime_of_jsval(datetime) {
            let _ = sheet.write_datetime_with_format(
                row,
                col,
                dt,
                &format.inner.lock().unwrap(),
            )?;
            Ok(self.clone())
        } else {
            Err(XlsxError::InvalidDate)
        }
    }

    #[wasm_bindgen(js_name = "writeUrlWithOptions", skip_jsdoc)]
    pub fn write_url_with_options(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        link: &Url,
        text: &str,
        tip: &str,
        format: Option<Format>,
    ) -> WasmResult<Worksheet> {
        let mut book = self.parent.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let _ = sheet.write_url_with_options(
            row,
            col,
            &*link.inner.lock().unwrap(),
            text,
            tip,
            format
                .map(|f| f.inner.lock().unwrap().clone())
                .as_ref(),
        )?;
        Ok(self.clone())
    }
}

// Each ConditionalFormat type gets its own method since wasm_bindgen exported
// structs don't implement JsCast (needed for runtime type dispatch via dyn_ref).
macro_rules! impl_add_cf {
    ($js_name:literal, $method:ident, $ty:ty) => {
        #[wasm_bindgen]
        impl Worksheet {
            #[wasm_bindgen(js_name = $js_name)]
            pub fn $method(
                &self,
                first_row: u32,
                first_col: u16,
                last_row: u32,
                last_col: u16,
                format: &$ty,
            ) -> WasmResult<Worksheet> {
                let mut book = self.parent.lock().unwrap();
                let sheet = book.worksheet_from_index(self.index).unwrap();
                sheet.add_conditional_format(
                    first_row, first_col, last_row, last_col,
                    &*format.inner.lock().unwrap(),
                )?;
                Ok(self.clone())
            }
        }
    };
}

impl_add_cf!("addConditionalFormat2ColorScale", add_conditional_format_2_color_scale, ConditionalFormat2ColorScale);
impl_add_cf!("addConditionalFormat3ColorScale", add_conditional_format_3_color_scale, ConditionalFormat3ColorScale);
impl_add_cf!("addConditionalFormatAverage", add_conditional_format_average, ConditionalFormatAverage);
impl_add_cf!("addConditionalFormatBlank", add_conditional_format_blank, ConditionalFormatBlank);
impl_add_cf!("addConditionalFormatCell", add_conditional_format_cell, ConditionalFormatCell);
impl_add_cf!("addConditionalFormatDataBar", add_conditional_format_data_bar, ConditionalFormatDataBar);
impl_add_cf!("addConditionalFormatDate", add_conditional_format_date, ConditionalFormatDate);
impl_add_cf!("addConditionalFormatDuplicate", add_conditional_format_duplicate, ConditionalFormatDuplicate);
impl_add_cf!("addConditionalFormatError", add_conditional_format_error, ConditionalFormatError);
impl_add_cf!("addConditionalFormatFormula", add_conditional_format_formula, ConditionalFormatFormula);
impl_add_cf!("addConditionalFormatIconSet", add_conditional_format_icon_set, ConditionalFormatIconSet);
impl_add_cf!("addConditionalFormatText", add_conditional_format_text, ConditionalFormatText);
impl_add_cf!("addConditionalFormatTop", add_conditional_format_top, ConditionalFormatTop);
