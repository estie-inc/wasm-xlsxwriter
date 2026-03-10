use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct DocProperties {
    pub(crate) inner: Arc<Mutex<xlsx::DocProperties>>,
}

#[wasm_bindgen]
impl DocProperties {
    #[wasm_bindgen(constructor)]
    pub fn new() -> DocProperties {
        DocProperties {
            inner: Arc::new(Mutex::new(xlsx::DocProperties::new())),
        }
    }
    #[wasm_bindgen(js_name = "setTitle", skip_jsdoc)]
    pub fn set_title(&self, title: &str) -> DocProperties {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_title(title);
        *lock = inner;
        DocProperties {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setSubject", skip_jsdoc)]
    pub fn set_subject(&self, subject: &str) -> DocProperties {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_subject(subject);
        *lock = inner;
        DocProperties {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setManager", skip_jsdoc)]
    pub fn set_manager(&self, manager: &str) -> DocProperties {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_manager(manager);
        *lock = inner;
        DocProperties {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setCompany", skip_jsdoc)]
    pub fn set_company(&self, company: &str) -> DocProperties {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_company(company);
        *lock = inner;
        DocProperties {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setCategory", skip_jsdoc)]
    pub fn set_category(&self, category: &str) -> DocProperties {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_category(category);
        *lock = inner;
        DocProperties {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setAuthor", skip_jsdoc)]
    pub fn set_author(&self, author: &str) -> DocProperties {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_author(author);
        *lock = inner;
        DocProperties {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setKeywords", skip_jsdoc)]
    pub fn set_keywords(&self, keywords: &str) -> DocProperties {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_keywords(keywords);
        *lock = inner;
        DocProperties {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setComment", skip_jsdoc)]
    pub fn set_comment(&self, comment: &str) -> DocProperties {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_comment(comment);
        *lock = inner;
        DocProperties {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setStatus", skip_jsdoc)]
    pub fn set_status(&self, status: &str) -> DocProperties {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_status(status);
        *lock = inner;
        DocProperties {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setHyperlinkBase", skip_jsdoc)]
    pub fn set_hyperlink_base(&self, hyperlink_base: &str) -> DocProperties {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_hyperlink_base(hyperlink_base);
        *lock = inner;
        DocProperties {
            inner: Arc::clone(&self.inner),
        }
    }
}
