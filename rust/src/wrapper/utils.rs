use chrono::DateTime;
use chrono::Utc;
use js_sys::Date;
use js_sys::Object;
use js_sys::Reflect;
use wasm_bindgen::convert::RefFromWasmAbi;
use wasm_bindgen::prelude::*;
use wasm_bindgen::JsCast;

use super::formula::Formula;
use super::url::Url;

fn is_jsdate(obj: &JsValue) -> bool {
    obj.is_instance_of::<Date>()
}

fn try_into_jsdate(obj: JsValue) -> Option<Date> {
    if is_jsdate(&obj) {
        Some(obj.unchecked_into())
    } else {
        None
    }
}

pub fn jsval_is_datetime(obj: &JsValue) -> bool {
    is_jsdate(obj)
}

pub fn datetime_of_jsval(obj: JsValue) -> Option<chrono::NaiveDateTime> {
    if is_jsdate(&obj) {
        let jsdate = try_into_jsdate(obj).unwrap();
        let timestamp_ms = jsdate.get_time() as i64;

        let seconds = timestamp_ms / 1000;
        let nanoseconds = (timestamp_ms % 1000) * 1_000_000;

        let dt = DateTime::<Utc>::from_timestamp(seconds, nanoseconds as u32);
        dt.map(|dt| dt.naive_utc())
    } else {
        None
    }
}

// https://github.com/rustwasm/wasm-bindgen/issues/2231#issuecomment-656293288
// wasm-pack 0.13.0 では、`ptr`ではなく`__wbg_ptr`にスタック上のポインタが格納される
pub fn generic_of_jsval<T: RefFromWasmAbi<Abi = u32>>(
    js: &JsValue,
    classname: &str,
) -> Result<T::Anchor, JsError> {
    if !js.is_object() {
        return Err(JsError::new(
            format!("Value supplied as {} is not an object", classname).as_str(),
        ));
    }

    let ctor_name = Object::get_prototype_of(js).constructor().name();
    if ctor_name == classname {
        let ptr = Reflect::get(js, &JsValue::from_str("__wbg_ptr")).map_err(|e| {
            JsError::new(format!("failed to get __wbg_ptr field: {:?}", e).as_str())
        })?;
        let ptr_u32: u32 = ptr.as_f64().ok_or(JsError::new(
            format!(
                "__wbg_ptr field of {} is not a number: {:?}",
                classname, ptr
            )
            .as_str(),
        ))? as u32;
        let data = unsafe { T::ref_from_abi(ptr_u32) };
        Ok(data)
    } else {
        Err(JsError::new(
            format!(
                "Value supplied as {} is not an instance of {}",
                classname, ctor_name
            )
            .as_str(),
        ))
    }
}

pub fn formula_of_jsval(js: &JsValue) -> Option<Formula> {
    generic_of_jsval::<Formula>(js, "Formula")
        .ok()
        .map(|f| f.clone())
}

pub fn url_of_jsval(js: &JsValue) -> Option<Url> {
    generic_of_jsval::<Url>(js, "Url").ok().map(|f| f.clone())
}
