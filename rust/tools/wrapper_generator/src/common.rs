use ruast::*;

/// Generate common use statements for struct wrappers
pub fn generate_use_statements() -> Vec<Use> {
    vec![
        // use rust_xlsxwriter as xlsx;
        Use::from(Path::single("rust_xlsxwriter").chain_use_rename("xlsx")),
        // use std::sync::{Arc, Mutex, MutexGuard}
        Use::from(Path::single("std").chain("sync").chain_use_group(vec![
            UseTree::from(Path::single(PathSegment::new("Arc", None))),
            UseTree::from(Path::single(PathSegment::new("Mutex", None))),
            UseTree::from(Path::single(PathSegment::new("MutexGuard", None))),
        ])),
        // use wasm_bindgen::prelude::*;
        Use::from(
            Path::single("wasm_bindgen")
                .chain("prelude")
                .chain_use_glob(),
        ),
        // use crate::impl_method;
        Use::from(Path::single("crate").chain("impl_method")),
    ]
}

/// Generate use statements for enum wrappers
pub fn generate_enum_use_statements() -> Vec<Use> {
    vec![
        // use rust_xlsxwriter as xlsx;
        Use::from(Path::single("rust_xlsxwriter").chain_use_rename("xlsx")),
        // use wasm_bindgen::prelude::*;
        Use::from(
            Path::single("wasm_bindgen")
                .chain("prelude")
                .chain_use_glob(),
        ),
    ]
}