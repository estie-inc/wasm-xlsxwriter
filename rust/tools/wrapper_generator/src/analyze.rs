// Parse upstream rustdoc-json (crate-inspector) and convert to IR

use std::collections::HashMap;

use crate::ir::{
    AnalyzedConstructor, AnalyzedCrate, AnalyzedEnum, AnalyzedMethod, AnalyzedParam,
    AnalyzedStruct, AnalyzedVariant, Accessor, MethodOverride, ParamType, ReceiverKind, ReturnKind,
    StructRole, VariantKind,
};
use crate::utils::{process_doc_comment, to_camel_case};
use crate_inspector::{CrateItem, EnumItem, FunctionItem, StructItem};
use rustdoc_types::{GenericArg, GenericArgs, GenericBound, GenericParamDefKind, WherePredicate};

/// Information about a child type accessed from a parent.
struct ParentInfo {
    parent_name: String,
    /// (parent method name, child type name)
    accessors: Vec<(String, String)>,
}

/// Scan all structs' public methods and record those returning `&mut OtherStruct`
/// as parent-to-child relationships.
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
                if !func.is_method() {
                    continue;
                }
                let Some(output) = func.output() else {
                    continue;
                };
                // Look for methods that return `&mut OtherStruct`
                if let rustdoc_types::Type::BorrowedRef {
                    is_mutable: true,
                    type_: inner,
                    ..
                } = output
                {
                    if let rustdoc_types::Type::ResolvedPath(path) = inner.as_ref() {
                        let child_name = path_to_short_name(&path.path).to_string();
                        if child_name == parent_name {
                            continue;
                        }
                        let method_name = func.name().to_string();
                        let entry = child_to_parent
                            .entry(child_name.clone())
                            .or_insert_with(|| ParentInfo {
                                parent_name: parent_name.clone(),
                                accessors: Vec::new(),
                            });
                        // Record multiple accessors that return the same child
                        if entry.parent_name == parent_name {
                            entry.accessors.push((method_name, child_name));
                        }
                    }
                }
            }
        }
    }

    child_to_parent
}

/// Return the last segment, e.g. `"rust_xlsxwriter::format::Format"` -> `"Format"`.
fn path_to_short_name(path: &str) -> &str {
    path.rsplit("::").next().unwrap_or(path)
}

pub fn analyze_crate(krate: &crate_inspector::Crate) -> AnalyzedCrate {
    let parent_child = build_parent_child_graph(krate);

    let structs: Vec<AnalyzedStruct> = krate
        .all_structs()
        .filter(|s| s.is_crate_item() && s.is_public())
        .filter_map(|s| analyze_struct(&s, &parent_child))
        .collect();

    let enums: Vec<AnalyzedEnum> = krate
        .all_enums()
        .filter(|e| e.is_crate_item() && e.is_public())
        .map(|e| analyze_enum(&e))
        .collect();

    AnalyzedCrate { structs, enums }
}

fn analyze_struct(
    struct_item: &StructItem<'_>,
    parent_child: &HashMap<String, ParentInfo>,
) -> Option<AnalyzedStruct> {
    let name = struct_item.name().to_string();

    let role = if let Some(info) = parent_child.get(&name) {
        let accessors = info
            .accessors
            .iter()
            .map(|(method_name, _)| Accessor {
                parent_method: method_name.clone(),
                js_name: to_camel_case(method_name),
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

    let doc = struct_item
        .item()
        .docs
        .as_deref()
        .map(process_doc_comment)
        .filter(|s| !s.is_empty());

    let mut constructor: Option<AnalyzedConstructor> = None;
    let mut methods: Vec<AnalyzedMethod> = Vec::new();

    for impl_item in struct_item.associated_impls() {
        for func in impl_item.functions() {
            let func_name = func.name();

            if func_name == "new" && !func.is_method() {
                constructor = Some(analyze_constructor(&func));
            } else if func.is_method() && func_name != "new" {
                if let Some(method) = analyze_method(&func, &name) {
                    methods.push(method);
                }
            }
        }
    }

    Some(AnalyzedStruct {
        name,
        role,
        has_default,
        constructor,
        methods,
        doc,
    })
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
                other => ParamType::WrappedType(other.to_string()),
            }
        }

        rustdoc_types::Type::ImplTrait(bounds) => {
            // `impl Into<T>` -> resolve to T
            if let Some(inner_type) = extract_into_inner_from_bounds(bounds) {
                resolve_param_type(inner_type, generics)
            } else {
                ParamType::Unknown(type_to_string(ty))
            }
        }

        rustdoc_types::Type::Generic(name) => {
            // Search generic parameters for an `Into<T>` bound
            if let Some(inner_type) = resolve_generic_into(name, generics) {
                resolve_param_type(inner_type, generics)
            } else {
                ParamType::Unknown(name.clone())
            }
        }

        other => ParamType::Unknown(type_to_string(other)),
    }
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

fn analyze_enum(enum_item: &EnumItem<'_>) -> AnalyzedEnum {
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
        .map(process_doc_comment)
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

    AnalyzedEnum {
        name,
        variants,
        has_default,
        doc,
    }
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
