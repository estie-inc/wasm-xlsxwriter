use anyhow::{Context, Result};
use std::collections::HashMap;
use std::fs;
use rustdoc_types::{Crate, Id, ItemEnum, Visibility, Type, FunctionSignature};
use serde_json;
use rustdoc_json;

#[derive(Debug, Clone)]
pub struct StructInfo {
    pub name: String,
    pub doc: Option<String>,
    pub methods: Vec<MethodInfo>,
    pub functions: Vec<FunctionInfo>,
}

#[derive(Debug, Clone)]
pub struct MethodInfo {
    pub method_name: String,
    pub doc: Option<String>,
    pub sig: Option<FunctionSignature>,
    pub struct_name: String,
}

#[derive(Debug, Clone)]
pub struct FunctionInfo {
    pub name: String,
    pub doc: Option<String>,
    pub sig: Option<FunctionSignature>,
    pub module_name: String,
}

#[derive(Debug)]
pub struct EnumInfo {
    pub name: String,
    pub doc: Option<String>,
}

#[derive(Debug)]
pub struct ExtractedItems {
    pub structs: Vec<StructInfo>,
    pub enums: Vec<EnumInfo>,
}

/// Use rustdoc-json crate and rustdoc-types to generate crate information
pub fn get_crate_info(manifest_path: &str) -> Result<Crate> {
    let json_path = rustdoc_json::Builder::default()
        .toolchain("nightly")
        .manifest_path(manifest_path)
        .document_private_items(true)
        .build()
        .context(format!("Failed to generate JSON documentation for {}", manifest_path))?;

    let json_string = fs::read_to_string(&json_path)
        .context(format!("Failed to read JSON file: {}", json_path.to_string_lossy()))?;
    let crate_info: Crate = serde_json::from_str(&json_string)
        .context(format!("Failed to parse JSON file: {}", json_path.to_string_lossy()))?;

    Ok(crate_info)
}

/// Extracts public structs, methods, functions, and enums from a Rust crate.
pub fn extract_crate_items(crate_info: &Crate) -> ExtractedItems {
    let mut structs: HashMap<Id, StructInfo> = HashMap::new();
    let mut enums: Vec<EnumInfo> = Vec::new();

    // Standard Rust methods to filter out
    let standard_methods = [
        "clone", "clone_to_uninit", "clone_into", "borrow", "borrow_mut",
        "type_id", "to_owned", "into", "try_into", "eq", "fmt", "hash",
        "write", "equivalent"
    ];

    // First pass: identify all structs and enums
    for (id, item) in &crate_info.index {
        if !matches!(item.visibility, Visibility::Public) {
            continue;
        }
        match &item.inner {
            ItemEnum::Struct(_) => {
                if let Some(struct_name) = &item.name {
                    structs.insert(
                        item.id,
                        StructInfo {
                            name: struct_name.clone(),
                            doc: item.docs.clone(),
                            methods: Vec::new(),
                            functions: Vec::new(),
                        },
                    );
                }
            },
            ItemEnum::Enum(_) => {
                if let Some(enum_name) = &item.name {
                    enums.push(EnumInfo {
                        name: enum_name.clone(),
                        doc: item.docs.clone(),
                    });
                }
            },
            _ => {}
        }
    }

    // Second pass: find implementations and extract methods
    for (_id, item) in &crate_info.index {
        match &item.inner {
            ItemEnum::Impl(impl_data) => {
                if let Type::ResolvedPath(path) = &impl_data.for_ {
                    if let Some(struct_info) = structs.get_mut(&path.id) {
                        for item_id in &impl_data.items {
                            if let Some(method_item) = crate_info.index.get(item_id) {
                                if let ItemEnum::Function(function_data) = &method_item.inner {
                                    // Check if this is a method (has self parameter)
                                    let is_method = function_data.sig.inputs.iter().any(|(name, _)| name == "self");

                                    if is_method {
                                        if let Some(method_name) = &method_item.name {
                                            if !standard_methods.contains(&method_name.as_str()) && 
                                               matches!(method_item.visibility, Visibility::Public) {

                                                struct_info.methods.push(MethodInfo {
                                                    struct_name: struct_info.name.clone(),
                                                    method_name: method_name.clone(),
                                                    doc: method_item.docs.clone(),
                                                    sig: Some(function_data.sig.clone()),
                                                });
                                            }
                                        }
                                    } else {
                                        if let Some(function_name) = &method_item.name {
                                            if matches!(method_item.visibility, Visibility::Public) {
                                                struct_info.functions.push(FunctionInfo {
                                                    name: function_name.clone(),
                                                    module_name: struct_info.name.clone(),
                                                    doc: method_item.docs.clone(),
                                                    sig: Some(function_data.sig.clone()),
                                                });
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        struct_info.methods.sort_by_key(|m| m.method_name.clone());
                        struct_info.functions.sort_by_key(|f| f.name.clone());
                    }
                }
            },
            _ => {}
        }
    }

    let mut structs_vec: Vec<_> = structs.into_values().collect();
    structs_vec.sort_by(|a, b| a.name.cmp(&b.name));

    let mut enums_vec = enums;
    enums_vec.sort_by(|a, b| a.name.cmp(&b.name));

    ExtractedItems {
        structs: structs_vec,
        enums: enums_vec,
    }
}
