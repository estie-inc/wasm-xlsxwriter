//! Glue code to make up for the lack of `rust_xlsxwriter::IntoExcelData` trait

use js_sys::wasm_bindgen;
use rust_xlsxwriter as xlsx;
use wasm_bindgen::{prelude::wasm_bindgen, JsValue};

use crate::error::XlsxError;

use super::{datetime::ExcelDateTime, formula::Formula, rich_string::RichString, url::Url, utils};

// We only export the ExcelData type since ExcelDataArray and ExcelDataMatrix are used for
#[wasm_bindgen(typescript_custom_section)]
const EXCEL_DATA: &'static str = r#"
/**
 *  Data type that can be written to Excel's cells.
 *  You can write data to cells via {@link Worksheet#write}.
 */
export type ExcelData = undefined | string | number | boolean | Date | Formula | Url | RichString;

type ExcelDataArray = ExcelData[];

type ExcelDataMatrix = ExcelData[][];
"#;

#[wasm_bindgen]
extern "C" {
    #[wasm_bindgen(typescript_type = "ExcelData")]
    pub type JsExcelData;

    #[wasm_bindgen(typescript_type = "ExcelDataArray")]
    pub type JsExcelDataArray;

    #[wasm_bindgen(typescript_type = "ExcelDataMatrix")]
    pub type JsExcelDataMatrix;
}

impl TryInto<ExcelData> for &JsExcelData {
    type Error = XlsxError;

    fn try_into(self) -> Result<ExcelData, Self::Error> {
        let jsvalue = JsValue::from(self);
        jsvalue.try_into()
    }
}

impl TryInto<Vec<ExcelData>> for &JsExcelDataArray {
    type Error = XlsxError;

    fn try_into(self) -> Result<Vec<ExcelData>, Self::Error> {
        if !self.is_array() {
            let js_type = self.js_typeof().as_string().unwrap();
            return Err(XlsxError::Type(format!(
                "Expected an array but found {}",
                js_type
            )));
        }

        let array = js_sys::Array::from(self);
        let mut vec = Vec::with_capacity(array.length() as usize);
        for i in 0..array.length() {
            let jsvalue = array.get(i);
            let excel_data: ExcelData = jsvalue.try_into()?;
            vec.push(excel_data);
        }
        Ok(vec)
    }
}

fn jsvalue_to_vec(jsvalue: JsValue) -> Result<Vec<ExcelData>, XlsxError> {
    if !jsvalue.is_array() {
        return Err(XlsxError::Type("Expected an array".to_string()));
    }

    let js_array = js_sys::Array::from(&jsvalue);
    let mut vec = Vec::with_capacity(js_array.length() as usize);
    for i in 0..js_array.length() {
        let jsvalue = js_array.get(i);
        let excel_data: ExcelData = jsvalue.try_into()?;
        vec.push(excel_data);
    }
    Ok(vec)
}

impl TryInto<Vec<Vec<ExcelData>>> for &JsExcelDataMatrix {
    type Error = XlsxError;

    fn try_into(self) -> Result<Vec<Vec<ExcelData>>, Self::Error> {
        if !self.is_array() {
            return Err(XlsxError::Type("Expected an array".to_string()));
        }

        let array = js_sys::Array::from(self);
        let mut vec = Vec::with_capacity(array.length() as usize);
        for i in 0..array.length() {
            let js_array = array.get(i);
            let excel_data_array: Vec<ExcelData> = jsvalue_to_vec(js_array)?;
            vec.push(excel_data_array);
        }
        Ok(vec)
    }
}

pub enum ExcelData {
    None,
    String(String),
    Number(f64),
    Bool(bool),
    DateTime(chrono::NaiveDateTime),
    ExcelDateTime(ExcelDateTime),
    Formula(Formula),
    Url(Url),
    RichString(RichString),
}

impl TryInto<ExcelData> for JsValue {
    type Error = XlsxError;

    fn try_into(self) -> Result<ExcelData, Self::Error> {
        let js_type = self.js_typeof().as_string().unwrap();

        // https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Operators/typeof#description
        match js_type.as_str() {
            "undefined" => Ok(ExcelData::None),
            "string" => Ok(ExcelData::String(self.as_string().unwrap())),
            "number" => Ok(ExcelData::Number(self.as_f64().unwrap())),
            "bigint" => {
                todo!("BigInt is not supported yet")
            }
            "boolean" => Ok(ExcelData::Bool(self.as_bool().unwrap())),
            "object" => {
                // FIXME: what should we do when writing `null`?
                if self.is_null() {
                    Ok(ExcelData::None)
                } else if utils::jsval_is_datetime(&self) {
                    let dt = utils::datetime_of_jsval(self).unwrap();
                    Ok(ExcelData::DateTime(dt))
                } else if let Some(excel_dt) = utils::excel_datetime_of_jsval(&self) {
                    Ok(ExcelData::ExcelDateTime(excel_dt.clone()))
                } else if let Some(formula) = utils::formula_of_jsval(&self) {
                    Ok(ExcelData::Formula(formula))
                } else if let Some(url) = utils::url_of_jsval(&self) {
                    Ok(ExcelData::Url(url))
                } else if let Some(rich_string) = utils::rich_string_of_jsval(&self) {
                    Ok(ExcelData::RichString(rich_string))
                } else {
                    let ctor = js_sys::Object::get_prototype_of(&self).constructor().name();
                    Err(XlsxError::Type(format!(
                        "Cannot write {} (instance of {}) to a cell",
                        js_type, ctor
                    )))
                }
            }
            "function" | "symbol" => {
                Err(XlsxError::Type(format!("Cannot write {js_type} to a cell")))
            }
            _ => unreachable!(),
        }
    }
}

impl xlsx::IntoExcelData for ExcelData {
    fn write(
        self,
        worksheet: &mut rust_xlsxwriter::Worksheet,
        row: rust_xlsxwriter::RowNum,
        col: rust_xlsxwriter::ColNum,
    ) -> Result<&mut rust_xlsxwriter::Worksheet, rust_xlsxwriter::XlsxError> {
        match self {
            ExcelData::None => worksheet.write_string(row, col, ""),
            ExcelData::String(s) => worksheet.write_string(row, col, &s),
            ExcelData::Number(n) => worksheet.write_number(row, col, n),
            ExcelData::Bool(b) => worksheet.write_boolean(row, col, b),
            ExcelData::DateTime(dt) => worksheet.write_datetime(row, col, dt),
            ExcelData::ExcelDateTime(excel_dt) => {
                let dt = excel_dt.inner.lock().unwrap().clone();
                worksheet.write_datetime(row, col, dt)
            }
            ExcelData::RichString(rich_string) => {
                let rich_string = rich_string.lock();
                let rich_string: Vec<_> =
                    rich_string.iter().map(|(f, s)| (f, s.as_str())).collect();
                worksheet.write_rich_string(row, col, &rich_string)
            }
            ExcelData::Formula(f) => worksheet.write_formula(row, col, &*f.lock()),
            ExcelData::Url(url) => worksheet.write_url(row, col, &*url.lock()),
        }
    }

    fn write_with_format<'a>(
        self,
        worksheet: &'a mut rust_xlsxwriter::Worksheet,
        row: rust_xlsxwriter::RowNum,
        col: rust_xlsxwriter::ColNum,
        format: &rust_xlsxwriter::Format,
    ) -> Result<&'a mut rust_xlsxwriter::Worksheet, rust_xlsxwriter::XlsxError> {
        match self {
            ExcelData::None => worksheet.write_blank(row, col, format),
            ExcelData::String(s) => worksheet.write_string_with_format(row, col, &s, format),
            ExcelData::Number(n) => worksheet.write_number_with_format(row, col, n, format),
            ExcelData::Bool(b) => worksheet.write_boolean_with_format(row, col, b, format),
            ExcelData::DateTime(dt) => worksheet.write_datetime_with_format(row, col, dt, format),
            ExcelData::ExcelDateTime(excel_dt) => {
                let dt = excel_dt.inner.lock().unwrap().clone();
                worksheet.write_datetime_with_format(row, col, dt, format)
            }
            ExcelData::Formula(f) => {
                worksheet.write_formula_with_format(row, col, &*f.lock(), format)
            }
            ExcelData::RichString(rich_string) => {
                let rich_string = rich_string.lock();
                let rich_string: Vec<_> =
                    rich_string.iter().map(|(f, s)| (f, s.as_str())).collect();
                worksheet.write_rich_string_with_format(row, col, &rich_string, format)
            }
            ExcelData::Url(url) => worksheet.write_url_with_format(row, col, &*url.lock(), format),
        }
    }
}
