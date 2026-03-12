use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `DocProperties` struct is used to create an object to represent document
/// metadata properties.
///
/// The `DocProperties` struct is used to create an object to represent various
/// document properties for an Excel file such as the Author's name or the
/// Creation Date.
///
/// Document Properties can be set for the "Summary" section and also for the
/// "Custom" section of the Excel document properties. See the examples below.
///
/// The `DocProperties` struct is used in conjunction with the
/// {@link Workbook#setProperties} method.
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
    #[doc = r" Create a deep clone of this object."]
    #[wasm_bindgen(js_name = "clone")]
    pub fn deep_clone(&self) -> DocProperties {
        DocProperties {
            inner: Arc::new(Mutex::new(self.inner.lock().unwrap().clone())),
        }
    }
    /// Set the Title field of the document properties.
    ///
    /// Set the "Title" field of the document properties to create a title for
    /// the document such as "Sales Report". See the example above.
    ///
    /// # Parameters
    ///
    /// - `title`: The title string property.
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
    /// Set the Subject field of the document properties.
    ///
    /// Set the "Subject" field of the document properties to indicate the
    /// subject matter. See the example above.
    ///
    /// # Parameters
    ///
    /// - `subject`: The subject string property.
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
    /// Set the Manager field of the document properties.
    ///
    /// Set the "Manager" field of the document properties. See the example
    /// above. See the example above.
    ///
    /// # Parameters
    ///
    /// - `manager`: The manager string property.
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
    /// Set the Company field of the document properties.
    ///
    /// Set the "Company" field of the document properties. See the example
    /// above.
    ///
    /// # Parameters
    ///
    /// - `company`: The company string property.
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
    /// Set the Category field of the document properties.
    ///
    /// Set the "Category" field of the document properties to indicate the
    /// category that the file belongs to. See the example above.
    ///
    /// # Parameters
    ///
    /// - `category`: The category string property.
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
    /// Set the Author field of the document properties.
    ///
    /// Set the "Author" field of the document properties. See the example
    /// above.
    ///
    /// # Parameters
    ///
    /// - `author`: The author string property.
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
    /// Set the Keywords field of the document properties.
    ///
    /// Set the "Keywords" field of the document properties. This can be one or
    /// more keywords that can be used in searches. See the example above.
    ///
    /// # Parameters
    ///
    /// - `keywords`: The keywords string property.
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
    /// Set the Comment field of the document properties.
    ///
    /// Set the "Comment" field of the document properties. This can be a
    /// general comment or summary that you want to add to the properties. See
    /// the example above.
    ///
    /// # Parameters
    ///
    /// - `comment`: The comment string property.
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
    /// Set the Status field of the document properties.
    ///
    /// Set the "Status" field of the document properties, such as "Draft" or
    /// "Final".
    ///
    /// # Parameters
    ///
    /// - `status`: The status string property.
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
    /// Set the hyperlink base field of the document properties.
    ///
    /// Set the "Hyperlink base" field of the document properties to have a
    /// default base URL.
    ///
    /// # Parameters
    ///
    /// - `hyperlink_base`: The hyperlink base string property.
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
