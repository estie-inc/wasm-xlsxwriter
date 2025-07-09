use std::sync::{Arc, Mutex};

use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

// The `ExcelDateTime` struct is used to represent an Excel date and/or time.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ExcelDateTime {
    pub(crate) inner: Arc<Mutex<xlsx::ExcelDateTime>>,
}

#[wasm_bindgen]
impl ExcelDateTime {
    // Create a `ExcelDateTime` instance from a string reference.
    #[wasm_bindgen(js_name = "parseFromStr")]
    pub fn parse_from_str(s: &str) -> Result<ExcelDateTime, JsValue> {
        let dt = xlsx::ExcelDateTime::parse_from_str(s)
            .map_err(|e| JsValue::from_str(&format!("{:?}", e)))?;
        Ok(ExcelDateTime {
            inner: Arc::new(Mutex::new(dt)),
        })
    }

    // Create a `ExcelDateTime` instance from years, months and days.
    #[wasm_bindgen(js_name = "fromYMD", constructor)]
    pub fn from_ymd(year: u16, month: u8, day: u8) -> Result<ExcelDateTime, JsValue> {
        let dt = xlsx::ExcelDateTime::from_ymd(year, month, day)
            .map_err(|e| JsValue::from_str(&format!("{:?}", e)))?;
        Ok(ExcelDateTime {
            inner: Arc::new(Mutex::new(dt)),
        })
    }

    // Create a `ExcelDateTime` instance from hours, minutes and seconds.
    #[wasm_bindgen(js_name = "fromHMS")]
    pub fn from_hms(hour: u16, min: u8, sec: f64) -> Result<ExcelDateTime, JsValue> {
        let dt = xlsx::ExcelDateTime::from_hms(hour, min, sec)
            .map_err(|e| JsValue::from_str(&format!("{:?}", e)))?;
        Ok(ExcelDateTime {
            inner: Arc::new(Mutex::new(dt)),
        })
    }

    // Create a `ExcelDateTime` instance from hours, minutes, seconds and milliseconds.
    #[wasm_bindgen(js_name = "fromHMSMilli")]
    pub fn from_hms_milli(hour: u16, min: u8, sec: u8, milli: u16) -> Result<ExcelDateTime, JsValue> {
        let dt = xlsx::ExcelDateTime::from_hms_milli(hour, min, sec, milli)
            .map_err(|e| JsValue::from_str(&format!("{:?}", e)))?;
        Ok(ExcelDateTime {
            inner: Arc::new(Mutex::new(dt)),
        })
    }

    // Adds to a `ExcelDateTime` date instance with hours, minutes and seconds.
    #[wasm_bindgen(js_name = "andHMS")]
    pub fn and_hms(&self, hour: u16, min: u8, sec: f64) -> Result<ExcelDateTime, JsValue> {
        let dt = self.inner.lock().unwrap().clone().and_hms(hour, min, sec)
            .map_err(|e| JsValue::from_str(&format!("{:?}", e)))?;
        Ok(ExcelDateTime {
            inner: Arc::new(Mutex::new(dt)),
        })
    }

    // Adds to a `ExcelDateTime` date instance with hours, minutes, seconds and milliseconds.
    #[wasm_bindgen(js_name = "andHMSMilli")]
    pub fn and_hms_milli(&self, hour: u16, min: u8, sec: u8, milli: u16) -> Result<ExcelDateTime, JsValue> {
        let dt = self.inner.lock().unwrap().clone().and_hms_milli(hour, min, sec, milli)
            .map_err(|e| JsValue::from_str(&format!("{:?}", e)))?;
        Ok(ExcelDateTime {
            inner: Arc::new(Mutex::new(dt)),
        })
    }

    // Create a `ExcelDateTime` instance from an Excel serial date.
    #[wasm_bindgen(js_name = "fromSerialDatetime")]
    pub fn from_serial_datetime(number: f64) -> Result<ExcelDateTime, JsValue> {
        let dt = xlsx::ExcelDateTime::from_serial_datetime(number)
            .map_err(|e| JsValue::from_str(&format!("{:?}", e)))?;
        Ok(ExcelDateTime {
            inner: Arc::new(Mutex::new(dt)),
        })
    }

    // Create a `ExcelDateTime` instance from a Unix time.
    #[wasm_bindgen(js_name = "fromTimestamp")]
    pub fn from_timestamp(timestamp: i64) -> Result<ExcelDateTime, JsValue> {
        let dt = xlsx::ExcelDateTime::from_timestamp(timestamp)
            .map_err(|e| JsValue::from_str(&format!("{:?}", e)))?;
        Ok(ExcelDateTime {
            inner: Arc::new(Mutex::new(dt)),
        })
    }

    // Convert the `ExcelDateTime` to an Excel serial date.
    #[wasm_bindgen(js_name = "toExcel")]
    pub fn to_excel(&self) -> f64 {
        self.inner.lock().unwrap().to_excel()
    }
}
