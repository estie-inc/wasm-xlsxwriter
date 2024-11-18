use std::sync::{Arc, Mutex};

use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

/// The `DocProperties` struct is used to create an object to represent document
/// metadata properties.
///
/// The `DocProperties` struct is used to create an object to represent various
/// document properties for an Excel file such as the Author's name or the
/// Creation Date.
///
/// <img src="https://rustxlsxwriter.github.io/images/app_doc_properties.png">
///
/// Document Properties can be set for the "Summary" section and also for the
/// "Custom" section of the Excel document properties. See the examples below.
///
/// The `DocProperties` struct is used in conjunction with the
/// {@link Workbook#setProperties} method.
///
/// TODO: example omitted
#[derive(Clone)]
#[wasm_bindgen]
pub struct DocProperties {
    pub(crate) inner: Arc<Mutex<xlsx::DocProperties>>,
}

#[wasm_bindgen]
impl DocProperties {
    pub(crate) fn lock(&self) -> std::sync::MutexGuard<'_, xlsx::DocProperties> {
        self.inner.lock().unwrap()
    }

    /// Create a new `DocProperties` class.
    #[wasm_bindgen(constructor)]
    pub fn new() -> DocProperties {
        DocProperties {
            inner: Arc::new(Mutex::new(xlsx::DocProperties::new())),
        }
    }

    /// Set the Title field of the document properties.
    ///
    /// Set the "Title" field of the document properties to create a title for
    /// the document such as "Sales Report". See the example above.
    ///
    /// @param {string} title - The title string property.
    /// @returns {DocProperties} - The DocProperties object.
    #[wasm_bindgen(js_name = "setTitle", skip_jsdoc)]
    pub fn set_title(&self, title: &str) -> DocProperties {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_title(title);
        let _ = std::mem::replace(&mut *lock, inner);
        DocProperties {
            inner: Arc::clone(&self.inner),
        }
    }

    /// Set the Subject field of the document properties.
    ///
    /// Set the "Subject" field of the document properties to indicate the
    /// subject matter. See the example above.
    ///
    /// @param {string} subject - The subject string property.
    /// @returns {DocProperties} - The DocProperties object.
    #[wasm_bindgen(js_name = "setSubject", skip_jsdoc)]
    pub fn set_subject(&self, subject: &str) -> DocProperties {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_subject(subject);
        let _ = std::mem::replace(&mut *lock, inner);
        DocProperties {
            inner: Arc::clone(&self.inner),
        }
    }

    /// Set the Author field of the document properties.
    ///
    /// Set the "Author" field of the document properties. See the example
    /// above.
    ///
    /// @param {string} author - The author string property.
    /// @returns {DocProperties} - The DocProperties object.
    #[wasm_bindgen(js_name = "setAuthor", skip_jsdoc)]
    pub fn set_author(&self, author: &str) -> DocProperties {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_author(author);
        let _ = std::mem::replace(&mut *lock, inner);
        DocProperties {
            inner: Arc::clone(&self.inner),
        }
    }

    /// Set the Manager field of the document properties.
    ///
    /// Set the "Manager" field of the document properties. See the example
    /// above. See the example above.
    ///
    /// @param {string} manager - The manager string property.
    /// @returns {DocProperties} - The DocProperties object.
    #[wasm_bindgen(js_name = "setManager", skip_jsdoc)]
    pub fn set_manager(&self, manager: &str) -> DocProperties {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_manager(manager);
        let _ = std::mem::replace(&mut *lock, inner);
        DocProperties {
            inner: Arc::clone(&self.inner),
        }
    }

    /// Set the Company field of the document properties.
    ///
    /// Set the "Company" field of the document properties. See the example
    /// above.
    ///
    /// @param {string} company - The company string property.
    /// @returns {DocProperties} - The DocProperties object.
    #[wasm_bindgen(js_name = "setCompany", skip_jsdoc)]
    pub fn set_company(&self, company: &str) -> DocProperties {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_company(company);
        let _ = std::mem::replace(&mut *lock, inner);
        DocProperties {
            inner: Arc::clone(&self.inner),
        }
    }

    /// Set the Category field of the document properties.
    ///
    /// Set the "Category" field of the document properties to indicate the
    /// category that the file belongs to. See the example above.
    ///
    /// @param {string} category - The category string property.
    /// @returns {DocProperties} - The DocProperties object.
    #[wasm_bindgen(js_name = "setCategory", skip_jsdoc)]
    pub fn set_category(&self, category: &str) -> DocProperties {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_category(category);
        let _ = std::mem::replace(&mut *lock, inner);
        DocProperties {
            inner: Arc::clone(&self.inner),
        }
    }

    /// Set the Keywords field of the document properties.
    ///
    /// Set the "Keywords" field of the document properties. This can be one or
    /// more keywords that can be used in searches. See the example above.
    ///
    /// @param {string} keywords - The keywords string property.
    /// @returns {DocProperties} - The DocProperties object.
    #[wasm_bindgen(js_name = "setKeywords", skip_jsdoc)]
    pub fn set_keywords(&self, keywords: &str) -> DocProperties {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_keywords(keywords);
        let _ = std::mem::replace(&mut *lock, inner);
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
    /// @param {string} comment - The comment string property.
    /// @returns {DocProperties} - The DocProperties object.
    #[wasm_bindgen(js_name = "setComment", skip_jsdoc)]
    pub fn set_comment(&self, comment: &str) -> DocProperties {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_comment(comment);
        let _ = std::mem::replace(&mut *lock, inner);
        DocProperties {
            inner: Arc::clone(&self.inner),
        }
    }

    /// Set the Status field of the document properties.
    ///
    /// Set the "Status" field of the document properties such as "Draft" or
    /// "Final".
    ///
    /// @param {string} status - The status string property.
    /// @returns {DocProperties} - The DocProperties object.
    #[wasm_bindgen(js_name = "setStatus", skip_jsdoc)]
    pub fn set_status(&self, status: &str) -> DocProperties {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_status(status);
        let _ = std::mem::replace(&mut *lock, inner);
        DocProperties {
            inner: Arc::clone(&self.inner),
        }
    }

    /// Set the hyperlink base field of the document properties.
    ///
    /// Set the "Hyperlink base" field of the document properties to have a
    /// default base url.
    ///
    /// @param {string} hyperlink_base - The hyperlink base string property.
    /// @returns {DocProperties} - The DocProperties object.
    #[wasm_bindgen(js_name = "setHyperlinkBase", skip_jsdoc)]
    pub fn set_hyperlink_base(&self, hyperlink_base: &str) -> DocProperties {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_hyperlink_base(hyperlink_base);
        let _ = std::mem::replace(&mut *lock, inner);
        DocProperties {
            inner: Arc::clone(&self.inner),
        }
    }
}
