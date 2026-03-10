// IR から struct ラッパーの Rust コードを生成する

use std::collections::HashSet;

use crate::ir::{AnalyzedMethod, AnalyzedParam, AnalyzedStruct, Accessor, ParamType, ReceiverKind, ReturnKind, StructRole};

/// Context for code generation.
pub struct CodegenContext {
    /// upstream crate の全 enum 名（.into() パターン判定用）
    pub enum_names: HashSet<String>,
    /// ラッパーに存在する型名（メソッド除外判定用）
    pub available_types: HashSet<String>,
    /// 手書きラッパーの struct 名（inner: T パターン。generated は inner: Arc<Mutex<T>>）
    pub handwritten_struct_names: HashSet<String>,
    /// 手書きラッパーで包まれた enum 名（inner: xlsx::T パターン。generated enum は From で変換）
    pub handwritten_enum_names: HashSet<String>,
}

impl CodegenContext {
    /// テスト用の空コンテキスト
    pub fn empty() -> Self {
        Self {
            enum_names: HashSet::new(),
            available_types: HashSet::new(),
            handwritten_struct_names: HashSet::new(),
            handwritten_enum_names: HashSet::new(),
        }
    }
}

pub fn generate_struct_file(s: &AnalyzedStruct, ctx: &CodegenContext) -> String {
    match &s.role {
        StructRole::Standalone => generate_standalone_struct(s, ctx),
        StructRole::Proxy { parent_name, accessors } => {
            generate_proxy_struct(s, parent_name, accessors, ctx)
        }
    }
}

fn collect_import_types(s: &AnalyzedStruct, available_types: &HashSet<String>) -> Vec<String> {
    let mut types: Vec<String> = vec![];
    let all_params = s
        .generatable_methods()
        .filter(|m| method_types_available(m, available_types))
        .flat_map(|m| m.params.iter())
        .chain(
            s.constructor
                .iter()
                .flat_map(|c| c.params.iter()),
        );
    for param in all_params {
        collect_wrapped_types_from_param(&param.ty, &mut types);
    }
    types.sort();
    types.dedup();
    types
}

fn method_types_available(m: &AnalyzedMethod, available: &HashSet<String>) -> bool {
    m.params.iter().all(|p| param_type_available(&p.ty, available))
}

fn param_type_available(ty: &ParamType, available: &HashSet<String>) -> bool {
    match ty {
        ParamType::WrappedType(name) | ParamType::RefWrappedType(name) => {
            available.contains(name)
        }
        ParamType::VecOf(inner) | ParamType::RefSliceOf(inner) | ParamType::OptionOf(inner) => {
            param_type_available(inner, available)
        }
        _ => true,
    }
}

fn collect_wrapped_types_from_param(ty: &ParamType, out: &mut Vec<String>) {
    match ty {
        ParamType::WrappedType(name) | ParamType::RefWrappedType(name) => {
            out.push(name.clone());
        }
        ParamType::VecOf(inner)
        | ParamType::RefSliceOf(inner)
        | ParamType::OptionOf(inner) => collect_wrapped_types_from_param(inner, out),
        _ => {}
    }
}

fn generate_standalone_struct(s: &AnalyzedStruct, ctx: &CodegenContext) -> String {
    let name = &s.name;
    let import_types = collect_import_types(s, &ctx.available_types);

    let mut lines: Vec<String> = vec![];

    lines.push("use rust_xlsxwriter as xlsx;".into());
    lines.push("use std::sync::{Arc, Mutex};".into());
    lines.push("use wasm_bindgen::prelude::*;".into());
    lines.push("use crate::wrapper::WasmResult;".into());
    for ty in &import_types {
        if ty != name {
            lines.push(format!("use crate::wrapper::{};", ty));
        }
    }
    lines.push(String::new());

    lines.push("#[derive(Clone)]".into());
    lines.push("#[wasm_bindgen]".into());
    lines.push(format!("pub struct {} {{", name));
    lines.push(format!("    pub(crate) inner: Arc<Mutex<xlsx::{}>>,", name));
    lines.push("}".into());
    lines.push(String::new());

    lines.push("#[wasm_bindgen]".into());
    lines.push(format!("impl {} {{", name));

    if let Some(ctor) = &s.constructor {
        let params_sig: Vec<String> = ctor.params.iter().map(format_param_sig).collect();
        let params_call: Vec<String> = ctor.params.iter().map(|p| format_param_call(p, ctx)).collect();
        lines.push("    #[wasm_bindgen(constructor)]".into());
        lines.push(format!(
            "    pub fn new({}) -> {} {{",
            params_sig.join(", "),
            name
        ));
        lines.push(format!("        {} {{", name));
        lines.push(format!(
            "            inner: Arc::new(Mutex::new(xlsx::{}::new({}))),",
            name,
            params_call.join(", ")
        ));
        lines.push("        }".into());
        lines.push("    }".into());
    }

    for method in s.generatable_methods() {
        if !method_types_available(method, &ctx.available_types) {
            continue;
        }
        let method_code = generate_method(method, name, s.has_default, ctx);
        lines.push(method_code);
    }

    lines.push("}".into());

    lines.join("\n")
}

fn generate_method(m: &AnalyzedMethod, struct_name: &str, has_default: bool, ctx: &CodegenContext) -> String {
    let params_sig: Vec<String> = m.params.iter().map(format_param_sig).collect();
    let params_call: Vec<String> = m.params.iter().map(|p| format_param_call(p, ctx)).collect();

    let ret_type = match &m.returns {
        ReturnKind::SelfType => struct_name.to_string(),
        ReturnKind::ResultSelf => format!("WasmResult<{}>", struct_name),
        ReturnKind::ResultVoid => "WasmResult<()>".to_string(),
        ReturnKind::Void => "()".to_string(),
        ReturnKind::Other(t) => t.clone(),
    };

    let sig_params = if params_sig.is_empty() {
        "&self".to_string()
    } else {
        format!("&self, {}", params_sig.join(", "))
    };

    let mut lines: Vec<String> = vec![];
    lines.push(format!(
        "    #[wasm_bindgen(js_name = \"{}\", skip_jsdoc)]",
        m.js_name
    ));
    lines.push(format!(
        "    pub fn {}({}) -> {} {{",
        m.name, sig_params, ret_type
    ));

    let call_args = params_call.join(", ");

    match (&m.returns, &m.receiver) {
        (ReturnKind::ResultSelf, _) => {
            lines.push("        let mut lock = self.inner.lock().unwrap();".into());
            lines.push(format!(
                "        lock.{}({})?;",
                m.name, call_args
            ));
            lines.push(format!(
                "        Ok({} {{ inner: Arc::clone(&self.inner) }})",
                struct_name
            ));
        }
        (ReturnKind::ResultVoid, _) => {
            lines.push("        let mut lock = self.inner.lock().unwrap();".into());
            lines.push(format!(
                "        lock.{}({})?;",
                m.name, call_args
            ));
            lines.push("        Ok(())".into());
        }
        (ReturnKind::Void, _) => {
            lines.push("        let mut lock = self.inner.lock().unwrap();".into());
            lines.push(format!("        lock.{}({});", m.name, call_args));
        }
        (ReturnKind::Other(ret), _) => {
            lines.push("        let lock = self.inner.lock().unwrap();".into());
            lines.push(format!("        lock.{}({})", m.name, call_args));
            let _ = ret;
        }
        (ReturnKind::SelfType, ReceiverKind::ConsumeSelf) => {
            if has_default {
                lines.push("        let mut lock = self.inner.lock().unwrap();".into());
                lines.push("        let mut inner = std::mem::take(&mut *lock);".into());
                lines.push(format!("        inner = inner.{}({});", m.name, call_args));
                lines.push("        *lock = inner;".into());
                lines.push(format!(
                    "        {} {{ inner: Arc::clone(&self.inner) }}",
                    struct_name
                ));
            } else {
                lines.push("        let mut lock = self.inner.lock().unwrap();".into());
                lines.push(format!(
                    "        let old = std::mem::replace(&mut *lock, xlsx::{}::default());",
                    struct_name
                ));
                lines.push(format!("        *lock = old.{}({});", m.name, call_args));
                lines.push(format!(
                    "        {} {{ inner: Arc::clone(&self.inner) }}",
                    struct_name
                ));
            }
        }
        (ReturnKind::SelfType, ReceiverKind::MutSelf) => {
            lines.push("        let mut lock = self.inner.lock().unwrap();".into());
            lines.push(format!("        lock.{}({});", m.name, call_args));
            lines.push(format!(
                "        {} {{ inner: Arc::clone(&self.inner) }}",
                struct_name
            ));
        }
        (ReturnKind::SelfType, ReceiverKind::RefSelf) => {
            lines.push("        let lock = self.inner.lock().unwrap();".into());
            lines.push(format!("        lock.{}({});", m.name, call_args));
            lines.push(format!(
                "        {} {{ inner: Arc::clone(&self.inner) }}",
                struct_name
            ));
        }
    }

    lines.push("    }".into());
    lines.join("\n")
}

fn generate_proxy_struct(s: &AnalyzedStruct, parent_name: &str, accessors: &[Accessor], ctx: &CodegenContext) -> String {
    let name = &s.name;
    let import_types = collect_import_types(s, &ctx.available_types);

    let multi_accessor = accessors.len() > 1;
    let accessor_enum_name = format!("{}Accessor", name);

    let mut lines: Vec<String> = vec![];

    lines.push("use rust_xlsxwriter as xlsx;".into());
    lines.push("use std::sync::{Arc, Mutex};".into());
    lines.push("use wasm_bindgen::prelude::*;".into());
    lines.push("use crate::wrapper::WasmResult;".into());
    for ty in &import_types {
        if ty != name {
            lines.push(format!("use crate::wrapper::{};", ty));
        }
    }
    lines.push(String::new());

    if multi_accessor {
        lines.push("#[derive(Clone, Copy)]".into());
        lines.push(format!("pub enum {} {{", accessor_enum_name));
        for acc in accessors {
            let variant = snake_to_pascal(&acc.parent_method);
            lines.push(format!("    {},", variant));
        }
        lines.push("}".into());
        lines.push(String::new());
    }

    lines.push("#[derive(Clone)]".into());
    lines.push("#[wasm_bindgen]".into());
    lines.push(format!("pub struct {} {{", name));
    lines.push(format!(
        "    pub(crate) parent: Arc<Mutex<xlsx::{}>>,",
        parent_name
    ));
    if multi_accessor {
        lines.push(format!(
            "    pub(crate) accessor: {},",
            accessor_enum_name
        ));
    }
    lines.push("}".into());
    lines.push(String::new());

    lines.push("#[wasm_bindgen]".into());
    lines.push(format!("impl {} {{", name));

    for method in s.generatable_methods() {
        if !method_types_available(method, &ctx.available_types) {
            continue;
        }
        let method_code =
            generate_proxy_method(method, name, parent_name, accessors, multi_accessor, ctx);
        lines.push(method_code);
    }

    lines.push("}".into());

    lines.join("\n")
}

fn generate_proxy_method(
    m: &AnalyzedMethod,
    struct_name: &str,
    _parent_name: &str,
    accessors: &[Accessor],
    multi_accessor: bool,
    ctx: &CodegenContext,
) -> String {
    let params_sig: Vec<String> = m.params.iter().map(format_param_sig).collect();
    let params_call: Vec<String> = m.params.iter().map(|p| format_param_call(p, ctx)).collect();

    let ret_type = match &m.returns {
        ReturnKind::SelfType => struct_name.to_string(),
        ReturnKind::ResultSelf => format!("WasmResult<{}>", struct_name),
        ReturnKind::ResultVoid => "WasmResult<()>".to_string(),
        ReturnKind::Void => "()".to_string(),
        ReturnKind::Other(t) => t.clone(),
    };

    let sig_params = if params_sig.is_empty() {
        "&self".to_string()
    } else {
        format!("&self, {}", params_sig.join(", "))
    };

    let call_args = params_call.join(", ");

    let mut lines: Vec<String> = vec![];
    lines.push(format!(
        "    #[wasm_bindgen(js_name = \"{}\", skip_jsdoc)]",
        m.js_name
    ));
    lines.push(format!(
        "    pub fn {}({}) -> {} {{",
        m.name, sig_params, ret_type
    ));
    lines.push("        let mut lock = self.parent.lock().unwrap();".into());

    let accessor_call = if multi_accessor {
        let accessor_enum_name = format!("{}Accessor", struct_name);
        // Build match expression for accessor
        let arms: Vec<String> = accessors
            .iter()
            .map(|acc| {
                let variant = snake_to_pascal(&acc.parent_method);
                format!(
                    "            {}::{} => lock.{}().{}({}),",
                    accessor_enum_name, variant, acc.parent_method, m.name, call_args
                )
            })
            .collect();
        format!(
            "        match self.accessor {{\n{}\n        }}",
            arms.join("\n")
        )
    } else {
        let acc = &accessors[0];
        format!(
            "        lock.{}().{}({});",
            acc.parent_method, m.name, call_args
        )
    };

    match &m.returns {
        ReturnKind::SelfType => {
            lines.push(accessor_call);
            lines.push(format!(
                "        {} {{ parent: Arc::clone(&self.parent) }}",
                struct_name
            ));
        }
        ReturnKind::ResultSelf => {
            // Treat accessor call as fallible
            lines.push(format!("        {}?;", accessor_call.trim_end_matches(';')));
            lines.push(format!(
                "        Ok({} {{ parent: Arc::clone(&self.parent) }})",
                struct_name
            ));
        }
        ReturnKind::ResultVoid => {
            lines.push(format!("        {}?;", accessor_call.trim_end_matches(';')));
            lines.push("        Ok(())".into());
        }
        ReturnKind::Void => {
            lines.push(accessor_call);
        }
        ReturnKind::Other(_) => {
            lines.push(accessor_call);
        }
    }

    lines.push("    }".into());
    lines.join("\n")
}

fn snake_to_pascal(s: &str) -> String {
    s.split('_')
        .map(|part| {
            let mut chars = part.chars();
            match chars.next() {
                None => String::new(),
                Some(first) => first.to_uppercase().collect::<String>() + chars.as_str(),
            }
        })
        .collect()
}

pub fn format_param_call(param: &AnalyzedParam, ctx: &CodegenContext) -> String {
    format_param_call_ty(&param.name, &param.ty, ctx)
}

/// struct パラメータの inner 値取得コード。
/// 手書きラッパー: inner: T → `.inner.clone()` / `&.inner`
/// 生成ラッパー: inner: Arc<Mutex<T>> → `.inner.lock().unwrap().clone()`
fn struct_inner_owned(name: &str, type_name: &str, ctx: &CodegenContext) -> String {
    if ctx.handwritten_struct_names.contains(type_name) {
        format!("{}.inner.clone()", name)
    } else {
        format!("{}.inner.lock().unwrap().clone()", name)
    }
}

fn struct_inner_ref(name: &str, type_name: &str, ctx: &CodegenContext) -> String {
    if ctx.handwritten_struct_names.contains(type_name) {
        format!("&{}.inner", name)
    } else {
        format!("&*{}.inner.lock().unwrap()", name)
    }
}

fn format_param_call_ty(name: &str, ty: &ParamType, ctx: &CodegenContext) -> String {
    match ty {
        ParamType::WrappedType(type_name) => {
            if ctx.handwritten_enum_names.contains(type_name) {
                format!("{}.inner", name)
            } else if ctx.enum_names.contains(type_name) {
                format!("xlsx::{}::from({})", type_name, name)
            } else {
                struct_inner_owned(name, type_name, ctx)
            }
        }
        ParamType::RefWrappedType(type_name) => {
            if ctx.handwritten_enum_names.contains(type_name) {
                format!("&{}.inner", name)
            } else if ctx.enum_names.contains(type_name) {
                format!("&xlsx::{}::from({})", type_name, name)
            } else {
                struct_inner_ref(name, type_name, ctx)
            }
        }
        ParamType::VecOf(inner) => match inner.as_ref() {
            ParamType::WrappedType(type_name) => {
                if ctx.handwritten_enum_names.contains(type_name) {
                    format!("{}.into_iter().map(|x| x.inner).collect()", name)
                } else if ctx.enum_names.contains(type_name) {
                    format!(
                        "{}.into_iter().map(|x| xlsx::{}::from(x)).collect()",
                        name, type_name
                    )
                } else if ctx.handwritten_struct_names.contains(type_name) {
                    format!("{}.iter().map(|x| x.inner.clone()).collect()", name)
                } else {
                    format!(
                        "{}.iter().map(|x| x.inner.lock().unwrap().clone()).collect()",
                        name
                    )
                }
            }
            _ => name.to_string(),
        },
        ParamType::RefSliceOf(inner) => match inner.as_ref() {
            ParamType::WrappedType(type_name) => {
                if ctx.handwritten_enum_names.contains(type_name) {
                    format!(
                        "&{}.iter().map(|x| x.inner).collect::<Vec<_>>()",
                        name
                    )
                } else if ctx.enum_names.contains(type_name) {
                    format!(
                        "&{}.into_iter().map(|x| xlsx::{}::from(x)).collect::<Vec<_>>()",
                        name, type_name
                    )
                } else if ctx.handwritten_struct_names.contains(type_name) {
                    format!(
                        "&{}.iter().map(|x| x.inner.clone()).collect::<Vec<_>>()",
                        name
                    )
                } else {
                    format!(
                        "&{}.iter().map(|x| x.inner.lock().unwrap().clone()).collect::<Vec<_>>()",
                        name
                    )
                }
            }
            _ => format!("&{}", name),
        },
        _ => name.to_string(),
    }
}

pub fn format_param_sig(param: &AnalyzedParam) -> String {
    format!("{}: {}", param.name, param.ty.to_rust_type_str())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ir::*;

    fn empty_enums() -> HashSet<String> {
        HashSet::new()
    }

    fn enum_set(names: &[&str]) -> HashSet<String> {
        names.iter().map(|s| s.to_string()).collect()
    }

    fn make_standalone(name: &str, has_default: bool) -> AnalyzedStruct {
        AnalyzedStruct {
            name: name.into(),
            role: StructRole::Standalone,
            has_default,
            constructor: Some(AnalyzedConstructor { params: vec![] }),
            methods: vec![],
            doc: None,
        }
    }

    #[test]
    fn standalone_struct_definition() {
        let s = make_standalone("Format", true);
        let code = generate_struct_file(&s, &CodegenContext::empty());
        assert!(code.contains("pub struct Format {"));
        assert!(code.contains("pub(crate) inner: Arc<Mutex<xlsx::Format>>,"));
        assert!(code.contains("#[derive(Clone)]"));
        assert!(code.contains("#[wasm_bindgen]"));
    }

    #[test]
    fn constructor_no_params() {
        let s = make_standalone("Format", true);
        let code = generate_struct_file(&s, &CodegenContext::empty());
        assert!(code.contains("#[wasm_bindgen(constructor)]"));
        assert!(code.contains("pub fn new() -> Format"));
        assert!(code.contains("Arc::new(Mutex::new(xlsx::Format::new()))"));
    }

    #[test]
    fn constructor_with_params() {
        let mut s = make_standalone("Note", false);
        s.constructor = Some(AnalyzedConstructor {
            params: vec![AnalyzedParam { name: "text".into(), ty: ParamType::Str }],
        });
        let code = generate_struct_file(&s, &CodegenContext::empty());
        assert!(code.contains("pub fn new(text: &str) -> Note"));
        assert!(code.contains("xlsx::Note::new(text)"));
    }

    #[test]
    fn consume_self_method_with_default() {
        let mut s = make_standalone("Format", true);
        s.methods.push(AnalyzedMethod {
            name: "set_bold".into(),
            js_name: "setBold".into(),
            receiver: ReceiverKind::ConsumeSelf,
            params: vec![],
            returns: ReturnKind::SelfType,
            override_: MethodOverride::Auto,
            doc: None,
        });
        let code = generate_struct_file(&s, &CodegenContext::empty());
        assert!(code.contains("pub fn set_bold(&self) -> Format"));
        assert!(code.contains("std::mem::take(&mut *lock)"));
        assert!(code.contains("inner = inner.set_bold()"));
        assert!(code.contains("Format { inner: Arc::clone(&self.inner) }"));
    }

    #[test]
    fn mut_self_method() {
        let mut s = make_standalone("ChartFont", true);
        s.methods.push(AnalyzedMethod {
            name: "set_bold".into(),
            js_name: "setBold".into(),
            receiver: ReceiverKind::MutSelf,
            params: vec![],
            returns: ReturnKind::SelfType,
            override_: MethodOverride::Auto,
            doc: None,
        });
        let code = generate_struct_file(&s, &CodegenContext::empty());
        assert!(code.contains("pub fn set_bold(&self) -> ChartFont"));
        assert!(code.contains("lock.set_bold()"));
        assert!(!code.contains("std::mem::take"));
    }

    #[test]
    fn method_with_primitive_param() {
        let mut s = make_standalone("Format", true);
        s.methods.push(AnalyzedMethod {
            name: "set_font_size".into(),
            js_name: "setFontSize".into(),
            receiver: ReceiverKind::ConsumeSelf,
            params: vec![AnalyzedParam { name: "size".into(), ty: ParamType::F64 }],
            returns: ReturnKind::SelfType,
            override_: MethodOverride::Auto,
            doc: None,
        });
        let code = generate_struct_file(&s, &CodegenContext::empty());
        assert!(code.contains("pub fn set_font_size(&self, size: f64) -> Format"));
        assert!(code.contains("inner.set_font_size(size)"));
    }

    #[test]
    fn method_with_enum_param() {
        let mut s = make_standalone("Format", true);
        s.methods.push(AnalyzedMethod {
            name: "set_align".into(),
            js_name: "setAlign".into(),
            receiver: ReceiverKind::ConsumeSelf,
            params: vec![AnalyzedParam {
                name: "align".into(),
                ty: ParamType::WrappedType("FormatAlign".into()),
            }],
            returns: ReturnKind::SelfType,
            override_: MethodOverride::Auto,
            doc: None,
        });
        let ctx = CodegenContext {
            enum_names: enum_set(&["FormatAlign"]),
            available_types: enum_set(&["FormatAlign"]),
            handwritten_struct_names: HashSet::new(),
            handwritten_enum_names: HashSet::new(),
        };
        let code = generate_struct_file(&s, &ctx);
        assert!(code.contains("pub fn set_align(&self, align: FormatAlign) -> Format"));
        assert!(code.contains("inner.set_align(xlsx::FormatAlign::from(align))"));
    }

    #[test]
    fn method_with_struct_param() {
        let mut s = make_standalone("ChartDataTable", true);
        s.methods.push(AnalyzedMethod {
            name: "set_font".into(),
            js_name: "setFont".into(),
            receiver: ReceiverKind::MutSelf,
            params: vec![AnalyzedParam {
                name: "font".into(),
                ty: ParamType::RefWrappedType("ChartFont".into()),
            }],
            returns: ReturnKind::SelfType,
            override_: MethodOverride::Auto,
            doc: None,
        });
        let ctx = CodegenContext {
            enum_names: HashSet::new(),
            available_types: enum_set(&["ChartFont"]),
            handwritten_struct_names: enum_set(&["ChartFont"]),
            handwritten_enum_names: HashSet::new(),
        };
        let code = generate_struct_file(&s, &ctx);
        assert!(code.contains("pub fn set_font(&self, font: ChartFont) -> ChartDataTable"));
        assert!(code.contains("&font.inner"));
    }

    #[test]
    fn result_returning_method() {
        let mut s = make_standalone("Worksheet", true);
        s.methods.push(AnalyzedMethod {
            name: "set_column_width".into(),
            js_name: "setColumnWidth".into(),
            receiver: ReceiverKind::MutSelf,
            params: vec![
                AnalyzedParam { name: "col".into(), ty: ParamType::U16 },
                AnalyzedParam { name: "width".into(), ty: ParamType::F64 },
            ],
            returns: ReturnKind::ResultSelf,
            override_: MethodOverride::Auto,
            doc: None,
        });
        let code = generate_struct_file(&s, &CodegenContext::empty());
        assert!(code.contains("-> WasmResult<Worksheet>"));
        assert!(code.contains("lock.set_column_width(col, width)?"));
        assert!(code.contains("Ok(Worksheet { inner: Arc::clone(&self.inner) })"));
    }

    #[test]
    fn skipped_method_not_generated() {
        let mut s = make_standalone("Chart", true);
        s.methods.push(AnalyzedMethod {
            name: "validate".into(),
            js_name: "validate".into(),
            receiver: ReceiverKind::RefSelf,
            params: vec![],
            returns: ReturnKind::Void,
            override_: MethodOverride::Skip("internal".into()),
            doc: None,
        });
        let code = generate_struct_file(&s, &CodegenContext::empty());
        assert!(!code.contains("fn validate"));
    }

    #[test]
    fn proxy_struct_single_accessor() {
        let s = AnalyzedStruct {
            name: "ChartTitle".into(),
            role: StructRole::Proxy {
                parent_name: "Chart".into(),
                accessors: vec![Accessor {
                    parent_method: "title".into(),
                    js_name: "title".into(),
                }],
            },
            has_default: false,
            constructor: None,
            methods: vec![AnalyzedMethod {
                name: "set_name".into(),
                js_name: "setName".into(),
                receiver: ReceiverKind::MutSelf,
                params: vec![AnalyzedParam { name: "name".into(), ty: ParamType::Str }],
                returns: ReturnKind::SelfType,
                override_: MethodOverride::Auto,
                doc: None,
            }],
            doc: None,
        };
        let code = generate_struct_file(&s, &CodegenContext::empty());
        assert!(code.contains("pub(crate) parent: Arc<Mutex<xlsx::Chart>>,"));
        assert!(code.contains("lock.title().set_name(name)"));
        assert!(code.contains("ChartTitle { parent: Arc::clone(&self.parent) }"));
    }
}
