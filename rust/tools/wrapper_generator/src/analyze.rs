// Parse upstream rustdoc-json (crate-inspector) and convert to IR

use std::collections::{HashMap, HashSet};

use crate::ir::{
    AnalyzedConstructor, AnalyzedCrate, AnalyzedEnum, AnalyzedMethod, AnalyzedParam,
    AnalyzedStruct, AnalyzedVariant, Accessor, MethodOverride, ParamType, ReceiverKind, ReturnKind,
    StructRole, VariantKind,
};

use crate::utils::{process_doc_comment, process_struct_doc_comment, to_camel_case};
use crate_inspector::{CrateItem, EnumItem, FunctionItem, StructItem};
use rustdoc_types::{GenericArg, GenericArgs, GenericBound, GenericParamDefKind, Id, WherePredicate};

/// Detect `#[doc(hidden)]` at the item level.
fn is_doc_hidden(item: &rustdoc_types::Item) -> bool {
    item.attrs.iter().any(|attr| {
        matches!(attr, rustdoc_types::Attribute::Other(s) if s.contains("doc(hidden)"))
    })
}

/// Build the set of module IDs that are publicly re-exported from the crate root
/// via `pub use module::*`. Structs and enums in these modules are part of the
/// crate's public API; others are internal types exposed only by `--document-hidden-items`.
fn build_public_module_ids(krate: &crate_inspector::Crate) -> HashSet<Id> {
    // krate.uses() already filters to root module use items
    krate
        .uses()
        .filter(|use_item| use_item.is_public() && use_item.is_glob())
        .filter_map(|use_item| use_item.inner().id)
        .collect()
}

/// Information about a child type accessed from a parent.
struct ParentInfo {
    parent_name: String,
    accessors: Vec<AccessorInfo>,
}

struct AccessorInfo {
    method_name: String,
    child_name: String,
    /// For indexed accessors (e.g., `worksheet_from_index(usize) -> Result<&mut T>`)
    key_type: Option<ParamType>,
}

/// Scan all structs' public methods and record those returning `&mut OtherStruct`
/// (or `Result<&mut OtherStruct, _>` with a key parameter) as parent-to-child relationships.
fn build_parent_child_graph(
    krate: &crate_inspector::Crate,
) -> HashMap<String, ParentInfo> {
    let mut child_to_parent: HashMap<String, ParentInfo> = HashMap::new();

    for struct_item in krate.all_structs() {
        if !struct_item.is_crate_item() {
            continue;
        }
        let parent_name = struct_item.name().to_string();

        for impl_item in struct_item.associated_impls() {
            for func in impl_item.functions() {
                if !func.is_method() || !func.is_public() {
                    continue;
                }
                let Some(output) = func.output() else {
                    continue;
                };

                // Try to extract the child type from the return type
                let detected = detect_proxy_accessor(output, &func, &parent_name);
                if let Some(info) = detected {
                    let entry = child_to_parent
                        .entry(info.child_name.clone())
                        .or_insert_with(|| ParentInfo {
                            parent_name: parent_name.clone(),
                            accessors: Vec::new(),
                        });
                    if entry.parent_name == parent_name {
                        entry.accessors.push(info);
                    }
                }
            }
        }
    }

    child_to_parent
}

/// Detect proxy accessor patterns:
/// 1. `fn accessor(&mut self) -> &mut ChildType` (standard proxy)
/// 2. `fn accessor(&mut self, key: primitive) -> Result<&mut ChildType, _>` (indexed proxy)
fn detect_proxy_accessor(
    output: &rustdoc_types::Type,
    func: &crate_inspector::FunctionItem<'_>,
    parent_name: &str,
) -> Option<AccessorInfo> {
    // Pattern 1: direct `&mut OtherStruct`
    if let Some(child_name) = extract_mut_ref_type(output) {
        if child_name != parent_name {
            return Some(AccessorInfo {
                method_name: func.name().to_string(),
                child_name,
                key_type: None,
            });
        }
    }

    // Pattern 2: `Result<&mut OtherStruct, _>` with a key parameter
    if let rustdoc_types::Type::ResolvedPath(path) = output {
        if path_to_short_name(&path.path) == "Result" {
            if let Some(args) = &path.args {
                if let GenericArgs::AngleBracketed { args: ab_args, .. } = args.as_ref() {
                    if let Some(GenericArg::Type(first_type)) = ab_args.first() {
                        if let Some(child_name) = extract_mut_ref_type(first_type) {
                            if child_name != parent_name {
                                // Check for a single non-self primitive parameter
                                if let Some(key_ty) = extract_single_primitive_param(func) {
                                    return Some(AccessorInfo {
                                        method_name: func.name().to_string(),
                                        child_name,
                                        key_type: Some(key_ty),
                                    });
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    None
}

fn extract_mut_ref_type(ty: &rustdoc_types::Type) -> Option<String> {
    if let rustdoc_types::Type::BorrowedRef {
        is_mutable: true,
        type_: inner,
        ..
    } = ty
    {
        if let rustdoc_types::Type::ResolvedPath(path) = inner.as_ref() {
            return Some(path_to_short_name(&path.path).to_string());
        }
    }
    None
}

/// If the method has exactly one non-self parameter and it's a primitive, return its ParamType.
fn extract_single_primitive_param(func: &crate_inspector::FunctionItem<'_>) -> Option<ParamType> {
    let inputs: Vec<_> = func.inputs().collect();
    let non_self: Vec<_> = inputs.iter().filter(|(name, _)| name != "self").collect();
    if non_self.len() != 1 {
        return None;
    }
    let (_, ty) = non_self[0];
    match ty {
        rustdoc_types::Type::Primitive(p) => match p.as_str() {
            "usize" => Some(ParamType::Usize),
            "u8" => Some(ParamType::U8),
            "u16" => Some(ParamType::U16),
            "u32" => Some(ParamType::U32),
            "u64" => Some(ParamType::U64),
            "i32" => Some(ParamType::I32),
            _ => None,
        },
        _ => None,
    }
}

/// Return the last segment, e.g. `"rust_xlsxwriter::format::Format"` -> `"Format"`.
fn path_to_short_name(path: &str) -> &str {
    path.rsplit("::").next().unwrap_or(path)
}

pub fn analyze_crate(krate: &crate_inspector::Crate) -> AnalyzedCrate {
    let parent_child = build_parent_child_graph(krate);
    let public_modules = build_public_module_ids(krate);

    let structs: Vec<AnalyzedStruct> = krate
        .all_structs()
        .filter(|s| s.is_crate_item() && s.is_public())
        .filter_map(|s| analyze_struct(&s, &parent_child, &public_modules))
        .collect();

    let enums: Vec<AnalyzedEnum> = krate
        .all_enums()
        .filter(|e| e.is_crate_item() && e.is_public())
        .filter_map(|e| analyze_enum(&e, &public_modules))
        .collect();

    AnalyzedCrate { structs, enums }
}

fn analyze_struct(
    struct_item: &StructItem<'_>,
    parent_child: &HashMap<String, ParentInfo>,
    public_modules: &HashSet<Id>,
) -> Option<AnalyzedStruct> {
    // Skip internal types: only process structs from publicly re-exported modules
    let in_public_module = struct_item
        .module()
        .map(|m| public_modules.contains(&m.item().id))
        .unwrap_or(false);
    if !in_public_module {
        return None;
    }
    // Also skip any items explicitly marked #[doc(hidden)]
    if is_doc_hidden(struct_item.item()) {
        return None;
    }

    let name = struct_item.name().to_string();

    let role = if let Some(info) = parent_child.get(&name) {
        let has_indexed = info.accessors.iter().any(|a| a.key_type.is_some());
        let accessors = info
            .accessors
            .iter()
            // When an indexed accessor exists, factory methods (add_xxx) are not accessors
            .filter(|acc| !has_indexed || acc.key_type.is_some())
            .map(|acc| Accessor {
                parent_method: acc.method_name.clone(),
                js_name: to_camel_case(&acc.method_name),
                key_type: acc.key_type.clone(),
            })
            .collect();
        StructRole::Proxy {
            parent_name: info.parent_name.clone(),
            accessors,
        }
    } else {
        StructRole::Standalone
    };

    let has_default = struct_item.trait_impls().any(|impl_item| {
        impl_item
            .trait_()
            .map(|t| path_to_short_name(&t.path) == "Default")
            .unwrap_or(false)
    });

    let has_clone = struct_item.trait_impls().any(|impl_item| {
        impl_item
            .trait_()
            .map(|t| path_to_short_name(&t.path) == "Clone")
            .unwrap_or(false)
    });

    let doc = struct_item
        .item()
        .docs
        .as_deref()
        .map(process_struct_doc_comment)
        .filter(|s| !s.is_empty());

    let mut constructor: Option<AnalyzedConstructor> = None;
    let mut methods: Vec<AnalyzedMethod> = Vec::new();

    for impl_item in struct_item.associated_impls() {
        for func in impl_item.functions() {
            // Skip non-public methods (pub(crate) etc. exposed by --document-hidden-items)
            if !func.is_public() {
                continue;
            }
            let func_name = func.name();

            if func_name == "new" && !func.is_method() {
                constructor = Some(analyze_constructor(&func));
            } else if func.is_method() && func_name != "new" && func_name != "deep_clone" {
                if let Some(method) = analyze_method(&func, &name) {
                    methods.push(method);
                }
            }
        }
    }

    // Auto-generate consume_self_default for types that:
    // - don't implement Default
    // - have a `new(...)` constructor with only zero-value-able params
    let consume_self_default = if !has_default {
        constructor
            .as_ref()
            .and_then(|ctor| auto_consume_self_default(&name, ctor))
    } else {
        None
    };

    Some(AnalyzedStruct {
        name,
        role,
        has_default,
        consume_self_default,
        constructor,
        methods,
        doc,
        has_clone,
    })
}

/// Generate a dummy constructor expression for types without Default that need
/// ConsumeSelf method support via `mem::replace`. Only succeeds when all constructor
/// params have known zero values; returns None if any param type is ambiguous.
fn auto_consume_self_default(name: &str, ctor: &AnalyzedConstructor) -> Option<String> {
    let mut args = Vec::new();
    for param in &ctor.params {
        match &param.ty {
            ParamType::Str => args.push(r#""""#.to_string()),
            ParamType::Bool => args.push("false".to_string()),
            ParamType::U8
            | ParamType::U16
            | ParamType::U32
            | ParamType::U64
            | ParamType::I8
            | ParamType::I16
            | ParamType::I32
            | ParamType::I64
            | ParamType::F32
            | ParamType::F64
            | ParamType::Usize => args.push("0".to_string()),
            _ => return None,
        }
    }
    Some(format!("xlsx::{}::new({})", name, args.join(", ")))
}

fn analyze_constructor(func: &FunctionItem<'_>) -> AnalyzedConstructor {
    let params = func
        .inputs()
        .filter(|(name, _)| name != "self")
        .map(|(name, ty)| AnalyzedParam {
            name: name.clone(),
            ty: resolve_param_type(ty, func.generics()),
        })
        .collect();

    AnalyzedConstructor { params }
}

fn analyze_method(func: &FunctionItem<'_>, struct_name: &str) -> Option<AnalyzedMethod> {
    let name = func.name().to_string();
    let js_name = to_camel_case(&name);
    let receiver = classify_receiver(func);
    let returns = classify_return(func, struct_name);

    let params: Vec<AnalyzedParam> = func
        .inputs()
        .filter(|(param_name, _)| param_name != "self")
        .map(|(param_name, ty)| AnalyzedParam {
            name: param_name.clone(),
            ty: resolve_param_type(ty, func.generics()),
        })
        .collect();

    let doc = func
        .item()
        .docs
        .as_deref()
        .map(process_doc_comment)
        .filter(|s| !s.is_empty());

    Some(AnalyzedMethod {
        name,
        js_name,
        receiver,
        params,
        returns,
        override_: MethodOverride::Auto,
        doc,
    })
}

fn classify_receiver(func: &FunctionItem<'_>) -> ReceiverKind {
    let self_param = func.inputs().find(|(name, _)| name == "self");

    match self_param {
        Some((_, rustdoc_types::Type::BorrowedRef { is_mutable: true, .. })) => {
            ReceiverKind::MutSelf
        }
        Some((_, rustdoc_types::Type::BorrowedRef { is_mutable: false, .. })) => {
            ReceiverKind::RefSelf
        }
        _ => ReceiverKind::ConsumeSelf,
    }
}

fn classify_return(func: &FunctionItem<'_>, struct_name: &str) -> ReturnKind {
    let Some(output) = func.output() else {
        return ReturnKind::Void;
    };

    match output {
        // `-> StructName` (owned self)
        rustdoc_types::Type::ResolvedPath(path) => {
            let short = path_to_short_name(&path.path);
            if short == struct_name {
                return ReturnKind::SelfType;
            }
            if short == "Result" {
                return classify_result_return(path, struct_name);
            }
            ReturnKind::Other(type_to_string(output))
        }
        // `-> &mut StructName`
        rustdoc_types::Type::BorrowedRef { type_: inner, .. } => {
            if let rustdoc_types::Type::ResolvedPath(path) = inner.as_ref() {
                let short = path_to_short_name(&path.path);
                if short == struct_name {
                    return ReturnKind::SelfType;
                }
            }
            ReturnKind::Other(type_to_string(output))
        }
        other => ReturnKind::Other(type_to_string(other)),
    }
}

/// Determine whether the return type is `Result<Self, E>` or `Result<(), E>`.
fn classify_result_return(
    result_path: &rustdoc_types::Path,
    struct_name: &str,
) -> ReturnKind {
    let Some(args_box) = &result_path.args else {
        return ReturnKind::Other("Result".to_string());
    };

    let GenericArgs::AngleBracketed { args, .. } = args_box.as_ref() else {
        return ReturnKind::Other("Result".to_string());
    };

    // Extract the first type argument of Result
    let first_type = args.iter().find_map(|arg| {
        if let GenericArg::Type(ty) = arg {
            Some(ty)
        } else {
            None
        }
    });

    match first_type {
        None => ReturnKind::Other("Result".to_string()),
        Some(rustdoc_types::Type::Tuple(fields)) if fields.is_empty() => ReturnKind::ResultVoid,
        Some(rustdoc_types::Type::ResolvedPath(inner_path)) => {
            let short = path_to_short_name(&inner_path.path);
            if short == struct_name {
                ReturnKind::ResultSelf
            } else {
                ReturnKind::Other(format!("Result<{}>", short))
            }
        }
        Some(rustdoc_types::Type::BorrowedRef { type_: inner, .. }) => {
            if let rustdoc_types::Type::ResolvedPath(inner_path) = inner.as_ref() {
                let short = path_to_short_name(&inner_path.path);
                if short == struct_name {
                    return ReturnKind::ResultSelf;
                }
            }
            ReturnKind::Other(format!("Result<{}>", type_to_string(first_type.unwrap())))
        }
        Some(other) => ReturnKind::Other(format!("Result<{}>", type_to_string(other))),
    }
}

pub(crate) fn resolve_param_type(
    ty: &rustdoc_types::Type,
    generics: &rustdoc_types::Generics,
) -> ParamType {
    match ty {
        rustdoc_types::Type::Primitive(name) => match name.as_str() {
            "bool" => ParamType::Bool,
            "u8" => ParamType::U8,
            "u16" => ParamType::U16,
            "u32" => ParamType::U32,
            "u64" => ParamType::U64,
            "i8" => ParamType::I8,
            "i16" => ParamType::I16,
            "i32" => ParamType::I32,
            "i64" => ParamType::I64,
            "f32" => ParamType::F32,
            "f64" => ParamType::F64,
            "usize" => ParamType::Usize,
            "str" => ParamType::Str,
            other => ParamType::Unknown(other.to_string()),
        },

        rustdoc_types::Type::BorrowedRef { type_: inner, .. } => {
            // Treat `&str` as Str
            if let rustdoc_types::Type::Primitive(name) = inner.as_ref() {
                if name == "str" {
                    return ParamType::Str;
                }
            }
            // `&[T]` -> RefSliceOf(T) -- upstream requires a slice reference
            if let rustdoc_types::Type::Slice(elem) = inner.as_ref() {
                let elem_type = resolve_param_type(elem, generics);
                return ParamType::RefSliceOf(Box::new(elem_type));
            }
            // `&WrappedType` -> RefWrappedType -- upstream requires a reference
            let inner_resolved = resolve_param_type(inner, generics);
            match inner_resolved {
                ParamType::WrappedType(name) => ParamType::RefWrappedType(name),
                other => other,
            }
        }

        rustdoc_types::Type::ResolvedPath(path) => {
            let short = path_to_short_name(&path.path);
            match short {
                "Vec" => {
                    if let Some(inner_type) = extract_first_type_arg(path) {
                        let elem_type = resolve_param_type(inner_type, generics);
                        ParamType::VecOf(Box::new(elem_type))
                    } else {
                        ParamType::Unknown("Vec<?>".to_string())
                    }
                }
                "Option" => {
                    if let Some(inner_type) = extract_first_type_arg(path) {
                        let inner = resolve_param_type(inner_type, generics);
                        ParamType::OptionOf(Box::new(inner))
                    } else {
                        ParamType::Unknown("Option<?>".to_string())
                    }
                }
                "String" => ParamType::Str,
                // Type aliases for primitive types used in rust_xlsxwriter
                "RowNum" => ParamType::U32,
                "ColNum" => ParamType::U16,
                other => ParamType::WrappedType(other.to_string()),
            }
        }

        rustdoc_types::Type::ImplTrait(bounds) => {
            // `impl Into<T>` -> resolve to T
            if let Some(inner_type) = extract_into_inner_from_bounds(bounds) {
                resolve_param_type(inner_type, generics)
            } else if has_asref_str_bound(bounds) {
                // `impl AsRef<str>` -> &str
                ParamType::Str
            } else if let Some(param_type) = extract_into_trait_type(bounds) {
                // `impl IntoChartFormat` -> MutRefWrappedType, `impl IntoChartRange` -> RefWrappedType
                param_type
            } else {
                ParamType::Unknown(type_to_string(ty))
            }
        }

        rustdoc_types::Type::Generic(name) => {
            // Search generic parameters for an `Into<T>` bound
            if let Some(inner_type) = resolve_generic_into(name, generics) {
                resolve_param_type(inner_type, generics)
            } else if let Some(param_type) = resolve_generic_into_trait(name, generics) {
                // T: IntoChartFormat -> MutRefWrappedType, T: IntoChartRange -> RefWrappedType
                param_type
            } else {
                ParamType::Unknown(name.clone())
            }
        }

        other => ParamType::Unknown(type_to_string(other)),
    }
}

/// Extract a ParamType from traits like `IntoChartRange` or `IntoChartFormat`.
/// `IntoChartFormat` maps to `MutRefWrappedType` because its impl is for `&mut ChartFormat`.
/// Other `IntoXxx` traits map to `RefWrappedType` (default immutable ref).
fn extract_into_trait_type(bounds: &[GenericBound]) -> Option<ParamType> {
    // Traits whose impls are for `&mut T` rather than `&T`
    const MUT_REF_TRAITS: &[&str] = &["IntoChartFormat"];

    for bound in bounds {
        if let GenericBound::TraitBound { trait_: path, .. } = bound {
            let short = path_to_short_name(&path.path);
            if let Some(rest) = short.strip_prefix("Into") {
                if !rest.is_empty() && rest.chars().next().is_some_and(|c| c.is_uppercase()) {
                    return if MUT_REF_TRAITS.contains(&short) {
                        Some(ParamType::MutRefWrappedType(rest.to_string()))
                    } else {
                        Some(ParamType::RefWrappedType(rest.to_string()))
                    };
                }
            }
        }
    }
    None
}

/// Search generic parameters or where clauses for `T: IntoXxx` and return the corresponding ParamType.
fn resolve_generic_into_trait(
    name: &str,
    generics: &rustdoc_types::Generics,
) -> Option<ParamType> {
    for param in &generics.params {
        if param.name != name {
            continue;
        }
        if let GenericParamDefKind::Type { bounds, .. } = &param.kind {
            if let Some(param_type) = extract_into_trait_type(bounds) {
                return Some(param_type);
            }
        }
    }

    for predicate in &generics.where_predicates {
        if let WherePredicate::BoundPredicate { type_, bounds, .. } = predicate {
            if let rustdoc_types::Type::Generic(pred_name) = type_ {
                if pred_name == name {
                    if let Some(param_type) = extract_into_trait_type(bounds) {
                        return Some(param_type);
                    }
                }
            }
        }
    }

    None
}

/// Check if bounds contain `AsRef<str>`.
fn has_asref_str_bound(bounds: &[GenericBound]) -> bool {
    for bound in bounds {
        if let GenericBound::TraitBound { trait_: path, .. } = bound {
            let short = path_to_short_name(&path.path);
            if short == "AsRef" {
                if let Some(inner) = extract_first_type_arg(path) {
                    if matches!(inner, rustdoc_types::Type::Primitive(s) if s == "str") {
                        return true;
                    }
                }
            }
        }
    }
    false
}

/// Extract T from a bounds list containing `Into<T>`.
fn extract_into_inner_from_bounds(bounds: &[GenericBound]) -> Option<&rustdoc_types::Type> {
    for bound in bounds {
        if let GenericBound::TraitBound { trait_: path, .. } = bound {
            let short = path_to_short_name(&path.path);
            if short == "Into" {
                if let Some(inner) = extract_first_type_arg(path) {
                    return Some(inner);
                }
            }
        }
    }
    None
}

/// Extract the first type argument from a `Path`'s angle-bracket arguments.
fn extract_first_type_arg(path: &rustdoc_types::Path) -> Option<&rustdoc_types::Type> {
    let args = path.args.as_deref()?;
    if let GenericArgs::AngleBracketed { args, .. } = args {
        for arg in args {
            if let GenericArg::Type(ty) = arg {
                return Some(ty);
            }
        }
    }
    None
}

/// Search generic parameters or where clauses for `T: Into<U>` and return U.
fn resolve_generic_into<'a>(
    name: &str,
    generics: &'a rustdoc_types::Generics,
) -> Option<&'a rustdoc_types::Type> {
    // First check bounds on the type parameter
    for param in &generics.params {
        if param.name != name {
            continue;
        }
        if let GenericParamDefKind::Type { bounds, .. } = &param.kind {
            if let Some(inner) = extract_into_inner_from_bounds(bounds) {
                return Some(inner);
            }
        }
    }

    // Check where clauses
    for predicate in &generics.where_predicates {
        if let WherePredicate::BoundPredicate { type_, bounds, .. } = predicate {
            if let rustdoc_types::Type::Generic(pred_name) = type_ {
                if pred_name == name {
                    if let Some(inner) = extract_into_inner_from_bounds(bounds) {
                        return Some(inner);
                    }
                }
            }
        }
    }

    None
}

fn analyze_enum(enum_item: &EnumItem<'_>, public_modules: &HashSet<Id>) -> Option<AnalyzedEnum> {
    // Skip internal types: only process enums from publicly re-exported modules
    let in_public_module = enum_item
        .module()
        .map(|m| public_modules.contains(&m.item().id))
        .unwrap_or(false);
    if !in_public_module {
        return None;
    }
    // Also skip any items explicitly marked #[doc(hidden)]
    if is_doc_hidden(enum_item.item()) {
        return None;
    }

    let name = enum_item.name().to_string();

    let has_default = enum_item.trait_impls().any(|impl_item| {
        impl_item
            .trait_()
            .map(|t| path_to_short_name(&t.path) == "Default")
            .unwrap_or(false)
    });

    let doc = enum_item
        .item()
        .docs
        .as_deref()
        .map(process_struct_doc_comment)
        .filter(|s| !s.is_empty());

    let variants = enum_item
        .variants()
        .map(|variant| {
            let variant_name = variant.name().to_string();
            let variant_doc = variant
                .item()
                .docs
                .as_deref()
                .map(process_doc_comment)
                .filter(|s| !s.is_empty());

            let kind = match variant.kind() {
                rustdoc_types::VariantKind::Plain => VariantKind::Plain,
                rustdoc_types::VariantKind::Tuple(fields) => {
                    let field_types: Vec<String> = fields
                        .iter()
                        .filter_map(|opt_id| opt_id.as_ref())
                        .filter_map(|id| {
                            let item = enum_item.krate().index.get(id)?;
                            if let rustdoc_types::ItemEnum::StructField(ty) = &item.inner {
                                Some(type_to_string(ty))
                            } else {
                                None
                            }
                        })
                        .collect();
                    VariantKind::Tuple(field_types)
                }
                rustdoc_types::VariantKind::Struct { fields, .. } => {
                    let field_pairs: Vec<(String, String)> = fields
                        .iter()
                        .filter_map(|id| {
                            let item = enum_item.krate().index.get(id)?;
                            let field_name = item.name.clone()?;
                            if let rustdoc_types::ItemEnum::StructField(ty) = &item.inner {
                                Some((field_name, type_to_string(ty)))
                            } else {
                                None
                            }
                        })
                        .collect();
                    VariantKind::Struct(field_pairs)
                }
            };

            AnalyzedVariant {
                name: variant_name,
                kind,
                doc: variant_doc,
            }
        })
        .collect();

    Some(AnalyzedEnum {
        name,
        variants,
        has_default,
        doc,
    })
}

/// Convert a type to a human-readable string (fallback for Unknown types).
fn type_to_string(ty: &rustdoc_types::Type) -> String {
    match ty {
        rustdoc_types::Type::Primitive(name) => name.clone(),
        rustdoc_types::Type::ResolvedPath(path) => {
            let short = path_to_short_name(&path.path);
            match &path.args {
                None => short.to_string(),
                Some(args_box) => match args_box.as_ref() {
                    GenericArgs::AngleBracketed { args, .. } => {
                        let arg_strs: Vec<String> = args
                            .iter()
                            .filter_map(|a| match a {
                                GenericArg::Type(t) => Some(type_to_string(t)),
                                GenericArg::Lifetime(l) => Some(l.clone()),
                                _ => None,
                            })
                            .collect();
                        if arg_strs.is_empty() {
                            short.to_string()
                        } else {
                            format!("{}<{}>", short, arg_strs.join(", "))
                        }
                    }
                    _ => short.to_string(),
                },
            }
        }
        rustdoc_types::Type::BorrowedRef {
            is_mutable,
            type_: inner,
            lifetime,
            ..
        } => {
            let lifetime_str = lifetime
                .as_deref()
                .map(|l| format!("{} ", l))
                .unwrap_or_default();
            let mut_str = if *is_mutable { "mut " } else { "" };
            format!("&{}{}{}", lifetime_str, mut_str, type_to_string(inner))
        }
        rustdoc_types::Type::Generic(name) => name.clone(),
        rustdoc_types::Type::Slice(inner) => format!("[{}]", type_to_string(inner)),
        rustdoc_types::Type::Tuple(fields) => {
            let field_strs: Vec<String> = fields.iter().map(type_to_string).collect();
            format!("({})", field_strs.join(", "))
        }
        rustdoc_types::Type::ImplTrait(bounds) => {
            let bound_strs: Vec<String> = bounds
                .iter()
                .filter_map(|b| {
                    if let GenericBound::TraitBound { trait_: path, .. } = b {
                        Some(path_to_short_name(&path.path).to_string())
                    } else {
                        None
                    }
                })
                .collect();
            format!("impl {}", bound_strs.join(" + "))
        }
        _ => "unknown".to_string(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use rustdoc_types::{GenericArgs, Generics, Id, Path, Type};

    fn empty_generics() -> Generics {
        Generics {
            params: vec![],
            where_predicates: vec![],
        }
    }

    fn make_path(name: &str) -> Path {
        Path {
            path: name.to_string(),
            id: Id(0),
            args: None,
        }
    }

    // ── resolve_param_type: primitives ─────────────────────────────────────────

    #[test]
    fn primitive_bool() {
        let ty = Type::Primitive("bool".to_string());
        assert_eq!(resolve_param_type(&ty, &empty_generics()), ParamType::Bool);
    }

    #[test]
    fn primitive_u32() {
        let ty = Type::Primitive("u32".to_string());
        assert_eq!(resolve_param_type(&ty, &empty_generics()), ParamType::U32);
    }

    #[test]
    fn primitive_i64() {
        let ty = Type::Primitive("i64".to_string());
        assert_eq!(resolve_param_type(&ty, &empty_generics()), ParamType::I64);
    }

    #[test]
    fn primitive_f64() {
        let ty = Type::Primitive("f64".to_string());
        assert_eq!(resolve_param_type(&ty, &empty_generics()), ParamType::F64);
    }

    #[test]
    fn primitive_str() {
        let ty = Type::Primitive("str".to_string());
        assert_eq!(resolve_param_type(&ty, &empty_generics()), ParamType::Str);
    }

    #[test]
    fn primitive_usize() {
        let ty = Type::Primitive("usize".to_string());
        assert_eq!(
            resolve_param_type(&ty, &empty_generics()),
            ParamType::Usize
        );
    }

    // ── resolve_param_type: borrowed ref ──────────────────────────────────────

    #[test]
    fn borrowed_str_ref() {
        let ty = Type::BorrowedRef {
            lifetime: None,
            is_mutable: false,
            type_: Box::new(Type::Primitive("str".to_string())),
        };
        assert_eq!(resolve_param_type(&ty, &empty_generics()), ParamType::Str);
    }

    #[test]
    fn borrowed_ref_to_struct_produces_ref_wrapped() {
        let ty = Type::BorrowedRef {
            lifetime: None,
            is_mutable: false,
            type_: Box::new(Type::ResolvedPath(make_path("Color"))),
        };
        assert_eq!(
            resolve_param_type(&ty, &empty_generics()),
            ParamType::RefWrappedType("Color".to_string())
        );
    }

    // ── resolve_param_type: ResolvedPath ──────────────────────────────────────

    #[test]
    fn resolved_path_color() {
        let ty = Type::ResolvedPath(make_path("Color"));
        assert_eq!(
            resolve_param_type(&ty, &empty_generics()),
            ParamType::WrappedType("Color".to_string())
        );
    }

    #[test]
    fn resolved_path_string_becomes_str() {
        let ty = Type::ResolvedPath(make_path("String"));
        assert_eq!(resolve_param_type(&ty, &empty_generics()), ParamType::Str);
    }

    #[test]
    fn resolved_path_vec_u8() {
        let inner_ty = Type::Primitive("u8".to_string());
        let path = Path {
            path: "Vec".to_string(),
            id: Id(0),
            args: Some(Box::new(GenericArgs::AngleBracketed {
                args: vec![GenericArg::Type(inner_ty)],
                constraints: vec![],
            })),
        };
        let ty = Type::ResolvedPath(path);
        assert_eq!(
            resolve_param_type(&ty, &empty_generics()),
            ParamType::VecOf(Box::new(ParamType::U8))
        );
    }

    #[test]
    fn resolved_path_option_bool() {
        let inner_ty = Type::Primitive("bool".to_string());
        let path = Path {
            path: "Option".to_string(),
            id: Id(0),
            args: Some(Box::new(GenericArgs::AngleBracketed {
                args: vec![GenericArg::Type(inner_ty)],
                constraints: vec![],
            })),
        };
        let ty = Type::ResolvedPath(path);
        assert_eq!(
            resolve_param_type(&ty, &empty_generics()),
            ParamType::OptionOf(Box::new(ParamType::Bool))
        );
    }

    // ── resolve_param_type: ImplTrait (impl Into<T>) ──────────────────────────

    #[test]
    fn impl_into_color() {
        let color_path = Path {
            path: "Color".to_string(),
            id: Id(0),
            args: None,
        };
        let into_path = Path {
            path: "Into".to_string(),
            id: Id(1),
            args: Some(Box::new(GenericArgs::AngleBracketed {
                args: vec![GenericArg::Type(Type::ResolvedPath(color_path))],
                constraints: vec![],
            })),
        };
        let bounds = vec![GenericBound::TraitBound {
            trait_: into_path,
            generic_params: vec![],
            modifier: rustdoc_types::TraitBoundModifier::None,
        }];
        let ty = Type::ImplTrait(bounds);
        assert_eq!(
            resolve_param_type(&ty, &empty_generics()),
            ParamType::WrappedType("Color".to_string())
        );
    }

    // ── type_to_string ────────────────────────────────────────────────────────

    #[test]
    fn type_to_string_primitive() {
        let ty = Type::Primitive("u32".to_string());
        assert_eq!(type_to_string(&ty), "u32");
    }

    #[test]
    fn type_to_string_borrowed_ref() {
        let ty = Type::BorrowedRef {
            lifetime: None,
            is_mutable: false,
            type_: Box::new(Type::Primitive("str".to_string())),
        };
        assert_eq!(type_to_string(&ty), "&str");
    }

    #[test]
    fn type_to_string_resolved_path_no_args() {
        let ty = Type::ResolvedPath(make_path("Format"));
        assert_eq!(type_to_string(&ty), "Format");
    }

    #[test]
    fn type_to_string_full_path_uses_short_name() {
        let ty = Type::ResolvedPath(make_path("rust_xlsxwriter::Format"));
        assert_eq!(type_to_string(&ty), "Format");
    }

    // ── path_to_short_name ────────────────────────────────────────────────────

    #[test]
    fn short_name_simple() {
        assert_eq!(path_to_short_name("Format"), "Format");
    }

    #[test]
    fn short_name_full_path() {
        assert_eq!(
            path_to_short_name("rust_xlsxwriter::format::Format"),
            "Format"
        );
    }

    // ── Integration test (requires external rust_xlsxwriter build) ────────────

    #[test]
    #[ignore]
    fn integration_analyze_rust_xlsxwriter() {
        use crate_inspector::CrateBuilder;
        use std::path::Path;

        let manifest = Path::new(
            "/Users/estie/estie/wasm-xlsxwriter/rust/tools/method_comparison/target/repos/rust_xlsxwriter/Cargo.toml",
        );

        let krate = CrateBuilder::default()
            .manifest_path(manifest)
            .build()
            .expect("failed to build rust_xlsxwriter crate");

        let analyzed = analyze_crate(&krate);

        // Verify Format struct is analyzed
        let format_struct = analyzed
            .structs
            .iter()
            .find(|s| s.name == "Format")
            .expect("Format struct not found");
        assert_eq!(format_struct.role, crate::ir::StructRole::Standalone);
        assert!(format_struct.has_default);
        assert!(format_struct.constructor.is_some());
        assert!(!format_struct.methods.is_empty());

        // Verify Color enum is analyzed
        let color_enum = analyzed
            .enums
            .iter()
            .find(|e| e.name == "Color")
            .expect("Color enum not found");
        assert!(!color_enum.variants.is_empty());

        // Verify ChartAxis is recognized as Proxy (Chart returns x_axis)
        if let Some(chart_axis) = analyzed.structs.iter().find(|s| s.name == "ChartAxis") {
            assert!(chart_axis.is_proxy());
        }
    }
}
