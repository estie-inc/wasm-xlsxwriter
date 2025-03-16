use std::sync::{Arc, Mutex};

use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;
use crate::macros::wrap_struct;

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

wrap_struct!(
    DocProperties,
    xlsx::DocProperties,
    set_title(title: &str),
    set_subject(subject: &str),
    set_author(author: &str),
    set_manager(manager: &str),
    set_company(company: &str),
    set_category(category: &str),
    set_keywords(keywords: &str),
    set_comment(comment: &str),
    set_status(status: &str),
    set_hyperlink_base(hyperlink_base: &str)
);
