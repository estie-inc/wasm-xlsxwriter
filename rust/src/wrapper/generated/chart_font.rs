use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartFont {
    pub(crate) inner: Arc<Mutex<xlsx::ChartFont>>,
}

#[wasm_bindgen]
impl ChartFont {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ChartFont {
        ChartFont {
            inner: Arc::new(Mutex::new(xlsx::ChartFont::new())),
        }
    }
    #[wasm_bindgen(js_name = "setBold", skip_jsdoc)]
    pub fn set_bold(&self) -> ChartFont {
        let mut lock = self.inner.lock().unwrap();
        lock.set_bold();
        ChartFont {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setItalic", skip_jsdoc)]
    pub fn set_italic(&self) -> ChartFont {
        let mut lock = self.inner.lock().unwrap();
        lock.set_italic();
        ChartFont {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setName", skip_jsdoc)]
    pub fn set_name(&self, font_name: &str) -> ChartFont {
        let mut lock = self.inner.lock().unwrap();
        lock.set_name(font_name);
        ChartFont {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setSize", skip_jsdoc)]
    pub fn set_size(&self, font_size: f64) -> ChartFont {
        let mut lock = self.inner.lock().unwrap();
        lock.set_size(font_size);
        ChartFont {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setRotation", skip_jsdoc)]
    pub fn set_rotation(&self, rotation: i16) -> ChartFont {
        let mut lock = self.inner.lock().unwrap();
        lock.set_rotation(rotation);
        ChartFont {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setUnderline", skip_jsdoc)]
    pub fn set_underline(&self) -> ChartFont {
        let mut lock = self.inner.lock().unwrap();
        lock.set_underline();
        ChartFont {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setStrikethrough", skip_jsdoc)]
    pub fn set_strikethrough(&self) -> ChartFont {
        let mut lock = self.inner.lock().unwrap();
        lock.set_strikethrough();
        ChartFont {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "unsetBold", skip_jsdoc)]
    pub fn unset_bold(&self) -> ChartFont {
        let mut lock = self.inner.lock().unwrap();
        lock.unset_bold();
        ChartFont {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setRightToLeft", skip_jsdoc)]
    pub fn set_right_to_left(&self, enable: bool) -> ChartFont {
        let mut lock = self.inner.lock().unwrap();
        lock.set_right_to_left(enable);
        ChartFont {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setPitchFamily", skip_jsdoc)]
    pub fn set_pitch_family(&self, family: u8) -> ChartFont {
        let mut lock = self.inner.lock().unwrap();
        lock.set_pitch_family(family);
        ChartFont {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setCharacterSet", skip_jsdoc)]
    pub fn set_character_set(&self, character_set: u8) -> ChartFont {
        let mut lock = self.inner.lock().unwrap();
        lock.set_character_set(character_set);
        ChartFont {
            inner: Arc::clone(&self.inner),
        }
    }
}
