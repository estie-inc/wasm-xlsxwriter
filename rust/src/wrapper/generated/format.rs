use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct Format {
    pub(crate) inner: Arc<Mutex<xlsx::Format>>,
}

#[wasm_bindgen]
impl Format {
    #[wasm_bindgen(constructor)]
    pub fn new() -> Format {
        Format {
            inner: Arc::new(Mutex::new(xlsx::Format::new())),
        }
    }
    #[wasm_bindgen(js_name = "merge", skip_jsdoc)]
    pub fn merge(&self, other: Format) -> Format {
        let lock = self.inner.lock().unwrap();
        lock.merge(&other.inner);
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setNumFormat", skip_jsdoc)]
    pub fn set_num_format(&self, num_format: &str) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_num_format(num_format);
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setNumFormatIndex", skip_jsdoc)]
    pub fn set_num_format_index(&self, num_format_index: u8) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_num_format_index(num_format_index);
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setBold", skip_jsdoc)]
    pub fn set_bold(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_bold();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setItalic", skip_jsdoc)]
    pub fn set_italic(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_italic();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setFontName", skip_jsdoc)]
    pub fn set_font_name(&self, font_name: &str) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_font_name(font_name);
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setFontSize", skip_jsdoc)]
    pub fn set_font_size(&self, font_size: f64) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_font_size(font_size);
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setFontFamily", skip_jsdoc)]
    pub fn set_font_family(&self, font_family: u8) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_font_family(font_family);
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setFontCharset", skip_jsdoc)]
    pub fn set_font_charset(&self, font_charset: u8) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_font_charset(font_charset);
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setFontStrikethrough", skip_jsdoc)]
    pub fn set_font_strikethrough(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_font_strikethrough();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setTextWrap", skip_jsdoc)]
    pub fn set_text_wrap(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_text_wrap();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setIndent", skip_jsdoc)]
    pub fn set_indent(&self, indent: u8) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_indent(indent);
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setRotation", skip_jsdoc)]
    pub fn set_rotation(&self, rotation: i16) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_rotation(rotation);
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setReadingDirection", skip_jsdoc)]
    pub fn set_reading_direction(&self, reading_direction: u8) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_reading_direction(reading_direction);
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setShrink", skip_jsdoc)]
    pub fn set_shrink(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_shrink();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setHyperlink", skip_jsdoc)]
    pub fn set_hyperlink(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_hyperlink();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setUnlocked", skip_jsdoc)]
    pub fn set_unlocked(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_unlocked();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setHidden", skip_jsdoc)]
    pub fn set_hidden(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_hidden();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setQuotePrefix", skip_jsdoc)]
    pub fn set_quote_prefix(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_quote_prefix();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setCheckbox", skip_jsdoc)]
    pub fn set_checkbox(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_checkbox();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "unsetBold", skip_jsdoc)]
    pub fn unset_bold(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.unset_bold();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "unsetItalic", skip_jsdoc)]
    pub fn unset_italic(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.unset_italic();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "unsetFontStrikethrough", skip_jsdoc)]
    pub fn unset_font_strikethrough(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.unset_font_strikethrough();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "unsetTextWrap", skip_jsdoc)]
    pub fn unset_text_wrap(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.unset_text_wrap();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "unsetShrink", skip_jsdoc)]
    pub fn unset_shrink(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.unset_shrink();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setLocked", skip_jsdoc)]
    pub fn set_locked(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_locked();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "unsetHidden", skip_jsdoc)]
    pub fn unset_hidden(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.unset_hidden();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "unsetHyperlinkStyle", skip_jsdoc)]
    pub fn unset_hyperlink_style(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.unset_hyperlink_style();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "unsetQuotePrefix", skip_jsdoc)]
    pub fn unset_quote_prefix(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.unset_quote_prefix();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "unsetCheckbox", skip_jsdoc)]
    pub fn unset_checkbox(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.unset_checkbox();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
}
