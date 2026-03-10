use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct ShapeFont {
    pub(crate) inner: Arc<Mutex<xlsx::ShapeFont>>,
}

#[wasm_bindgen]
impl ShapeFont {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ShapeFont {
        ShapeFont {
            inner: Arc::new(Mutex::new(xlsx::ShapeFont::new())),
        }
    }
    #[wasm_bindgen(js_name = "setBold", skip_jsdoc)]
    pub fn set_bold(&self) -> ShapeFont {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_bold();
        *lock = inner;
        ShapeFont {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setItalic", skip_jsdoc)]
    pub fn set_italic(&self) -> ShapeFont {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_italic();
        *lock = inner;
        ShapeFont {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setName", skip_jsdoc)]
    pub fn set_name(&self, font_name: &str) -> ShapeFont {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_name(font_name);
        *lock = inner;
        ShapeFont {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setSize", skip_jsdoc)]
    pub fn set_size(&self, font_size: f64) -> ShapeFont {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_size(font_size);
        *lock = inner;
        ShapeFont {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setUnderline", skip_jsdoc)]
    pub fn set_underline(&self) -> ShapeFont {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_underline();
        *lock = inner;
        ShapeFont {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setStrikethrough", skip_jsdoc)]
    pub fn set_strikethrough(&self) -> ShapeFont {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_strikethrough();
        *lock = inner;
        ShapeFont {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "unsetBold", skip_jsdoc)]
    pub fn unset_bold(&self) -> ShapeFont {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.unset_bold();
        *lock = inner;
        ShapeFont {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setRightToLeft", skip_jsdoc)]
    pub fn set_right_to_left(&self, enable: bool) -> ShapeFont {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_right_to_left(enable);
        *lock = inner;
        ShapeFont {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setPitchFamily", skip_jsdoc)]
    pub fn set_pitch_family(&self, family: u8) -> ShapeFont {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_pitch_family(family);
        *lock = inner;
        ShapeFont {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setCharacterSet", skip_jsdoc)]
    pub fn set_character_set(&self, character_set: u8) -> ShapeFont {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_character_set(character_set);
        *lock = inner;
        ShapeFont {
            inner: Arc::clone(&self.inner),
        }
    }
}
