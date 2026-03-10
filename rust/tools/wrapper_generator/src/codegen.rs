// IR から struct ラッパーの Rust コードを生成する

use std::collections::HashSet;

use proc_macro2::TokenStream;
use quote::{format_ident, quote};

use crate::codegen_tokens::{param_call_tokens, param_sig_tokens};
use crate::ir::{
    Accessor, AnalyzedMethod, AnalyzedStruct, ParamType, ReceiverKind, ReturnKind, StructRole,
};

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
        StructRole::Proxy {
            parent_name,
            accessors,
        } => generate_proxy_struct(s, parent_name, accessors, ctx),
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

fn generate_imports(name: &str, import_types: &[String]) -> TokenStream {
    let type_imports: Vec<TokenStream> = import_types
        .iter()
        .filter(|ty| ty.as_str() != name)
        .map(|ty| {
            let ty_ident = format_ident!("{}", ty);
            quote! { use crate::wrapper::#ty_ident; }
        })
        .collect();

    quote! {
        use rust_xlsxwriter as xlsx;
        use std::sync::{Arc, Mutex};
        use wasm_bindgen::prelude::*;
        use crate::wrapper::WasmResult;
        #(#type_imports)*
    }
}

fn generate_standalone_struct(s: &AnalyzedStruct, ctx: &CodegenContext) -> String {
    let name = &s.name;
    let name_ident = format_ident!("{}", name);
    let import_types = collect_import_types(s, &ctx.available_types);

    let imports = generate_imports(name, &import_types);

    let xlsx_name = format_ident!("{}", name);

    let ctor_tokens = if let Some(ctor) = &s.constructor {
        let params_sig: Vec<TokenStream> = ctor
            .params
            .iter()
            .map(|p| param_sig_tokens(&p.name, &p.ty))
            .collect();
        let params_call: Vec<TokenStream> = ctor
            .params
            .iter()
            .map(|p| param_call_tokens(&p.name, &p.ty, ctx))
            .collect();
        quote! {
            #[wasm_bindgen(constructor)]
            pub fn new(#(#params_sig),*) -> #name_ident {
                #name_ident {
                    inner: Arc::new(Mutex::new(xlsx::#xlsx_name::new(#(#params_call),*))),
                }
            }
        }
    } else {
        quote! {}
    };

    let method_tokens: Vec<TokenStream> = s
        .generatable_methods()
        .filter(|m| method_types_available(m, &ctx.available_types))
        .map(|m| generate_method(m, name, s.has_default, ctx))
        .collect();

    let struct_def = quote! {
        #[derive(Clone)]
        #[wasm_bindgen]
        pub struct #name_ident {
            pub(crate) inner: Arc<Mutex<xlsx::#xlsx_name>>,
        }
    };

    let impl_block = quote! {
        #[wasm_bindgen]
        impl #name_ident {
            #ctor_tokens
            #(#method_tokens)*
        }
    };

    format!("{}\n\n{}\n\n{}\n", imports, struct_def, impl_block)
}

fn generate_method(
    m: &AnalyzedMethod,
    struct_name: &str,
    has_default: bool,
    ctx: &CodegenContext,
) -> TokenStream {
    let struct_ident = format_ident!("{}", struct_name);
    let method_ident = format_ident!("{}", m.name);
    let js_name = &m.js_name;

    let params_sig: Vec<TokenStream> = m
        .params
        .iter()
        .map(|p| param_sig_tokens(&p.name, &p.ty))
        .collect();
    let params_call: Vec<TokenStream> = m
        .params
        .iter()
        .map(|p| param_call_tokens(&p.name, &p.ty, ctx))
        .collect();

    let ret_type: TokenStream = match &m.returns {
        ReturnKind::SelfType => quote! { #struct_ident },
        ReturnKind::ResultSelf => quote! { WasmResult<#struct_ident> },
        ReturnKind::ResultVoid => quote! { WasmResult<()> },
        ReturnKind::Void => quote! { () },
        ReturnKind::Other(t) => {
            let t_tokens: TokenStream = t.parse().expect("return type should be valid tokens");
            quote! { #t_tokens }
        }
    };

    let sig_params = if params_sig.is_empty() {
        quote! { &self }
    } else {
        quote! { &self, #(#params_sig),* }
    };

    let body: TokenStream = match (&m.returns, &m.receiver) {
        (ReturnKind::ResultSelf, _) => {
            quote! {
                let mut lock = self.inner.lock().unwrap();
                lock.#method_ident(#(#params_call),*)?;
                Ok(#struct_ident { inner: Arc::clone(&self.inner) })
            }
        }
        (ReturnKind::ResultVoid, _) => {
            quote! {
                let mut lock = self.inner.lock().unwrap();
                lock.#method_ident(#(#params_call),*)?;
                Ok(())
            }
        }
        (ReturnKind::Void, _) => {
            quote! {
                let mut lock = self.inner.lock().unwrap();
                lock.#method_ident(#(#params_call),*);
            }
        }
        (ReturnKind::Other(_), _) => {
            quote! {
                let lock = self.inner.lock().unwrap();
                lock.#method_ident(#(#params_call),*)
            }
        }
        (ReturnKind::SelfType, ReceiverKind::ConsumeSelf) => {
            let xlsx_ident = format_ident!("{}", struct_name);
            if has_default {
                quote! {
                    let mut lock = self.inner.lock().unwrap();
                    let mut inner = std::mem::take(&mut *lock);
                    inner = inner.#method_ident(#(#params_call),*);
                    *lock = inner;
                    #struct_ident { inner: Arc::clone(&self.inner) }
                }
            } else {
                quote! {
                    let mut lock = self.inner.lock().unwrap();
                    let old = std::mem::replace(&mut *lock, xlsx::#xlsx_ident::default());
                    *lock = old.#method_ident(#(#params_call),*);
                    #struct_ident { inner: Arc::clone(&self.inner) }
                }
            }
        }
        (ReturnKind::SelfType, ReceiverKind::MutSelf) => {
            quote! {
                let mut lock = self.inner.lock().unwrap();
                lock.#method_ident(#(#params_call),*);
                #struct_ident { inner: Arc::clone(&self.inner) }
            }
        }
        (ReturnKind::SelfType, ReceiverKind::RefSelf) => {
            quote! {
                let lock = self.inner.lock().unwrap();
                lock.#method_ident(#(#params_call),*);
                #struct_ident { inner: Arc::clone(&self.inner) }
            }
        }
    };

    quote! {
        #[wasm_bindgen(js_name = #js_name, skip_jsdoc)]
        pub fn #method_ident(#sig_params) -> #ret_type {
            #body
        }
    }
}

fn generate_proxy_struct(
    s: &AnalyzedStruct,
    parent_name: &str,
    accessors: &[Accessor],
    ctx: &CodegenContext,
) -> String {
    let name = &s.name;
    let name_ident = format_ident!("{}", name);
    let parent_ident = format_ident!("{}", parent_name);
    let import_types = collect_import_types(s, &ctx.available_types);

    let multi_accessor = accessors.len() > 1;

    let imports = generate_imports(name, &import_types);

    let mut sections: Vec<String> = vec![imports.to_string()];

    if multi_accessor {
        let accessor_enum_ident = format_ident!("{}Accessor", name);
        let variants: Vec<proc_macro2::Ident> = accessors
            .iter()
            .map(|acc| format_ident!("{}", snake_to_pascal(&acc.parent_method)))
            .collect();
        sections.push(
            quote! {
                #[derive(Clone, Copy)]
                pub enum #accessor_enum_ident {
                    #(#variants),*
                }
            }
            .to_string(),
        );
    }

    let accessor_field = if multi_accessor {
        let accessor_enum_ident = format_ident!("{}Accessor", name);
        quote! {
            pub(crate) accessor: #accessor_enum_ident,
        }
    } else {
        quote! {}
    };

    let method_tokens: Vec<TokenStream> = s
        .generatable_methods()
        .filter(|m| method_types_available(m, &ctx.available_types))
        .map(|m| generate_proxy_method(m, name, parent_name, accessors, multi_accessor, ctx))
        .collect();

    sections.push(
        quote! {
            #[derive(Clone)]
            #[wasm_bindgen]
            pub struct #name_ident {
                pub(crate) parent: Arc<Mutex<xlsx::#parent_ident>>,
                #accessor_field
            }
        }
        .to_string(),
    );

    sections.push(
        quote! {
            #[wasm_bindgen]
            impl #name_ident {
                #(#method_tokens)*
            }
        }
        .to_string(),
    );

    sections.join("\n\n")
}

fn generate_proxy_method(
    m: &AnalyzedMethod,
    struct_name: &str,
    _parent_name: &str,
    accessors: &[Accessor],
    multi_accessor: bool,
    ctx: &CodegenContext,
) -> TokenStream {
    let struct_ident = format_ident!("{}", struct_name);
    let method_ident = format_ident!("{}", m.name);
    let js_name = &m.js_name;

    let params_sig: Vec<TokenStream> = m
        .params
        .iter()
        .map(|p| param_sig_tokens(&p.name, &p.ty))
        .collect();
    let params_call: Vec<TokenStream> = m
        .params
        .iter()
        .map(|p| param_call_tokens(&p.name, &p.ty, ctx))
        .collect();

    let ret_type: TokenStream = match &m.returns {
        ReturnKind::SelfType => quote! { #struct_ident },
        ReturnKind::ResultSelf => quote! { WasmResult<#struct_ident> },
        ReturnKind::ResultVoid => quote! { WasmResult<()> },
        ReturnKind::Void => quote! { () },
        ReturnKind::Other(t) => {
            let t_tokens: TokenStream = t.parse().expect("return type should be valid tokens");
            quote! { #t_tokens }
        }
    };

    let sig_params = if params_sig.is_empty() {
        quote! { &self }
    } else {
        quote! { &self, #(#params_sig),* }
    };

    let accessor_expr: TokenStream = if multi_accessor {
        let accessor_enum_ident = format_ident!("{}Accessor", struct_name);
        let arms: Vec<TokenStream> = accessors
            .iter()
            .map(|acc| {
                let variant = format_ident!("{}", snake_to_pascal(&acc.parent_method));
                let accessor_method = format_ident!("{}", acc.parent_method);
                quote! {
                    #accessor_enum_ident::#variant => lock.#accessor_method().#method_ident(#(#params_call),*)
                }
            })
            .collect();
        quote! {
            match self.accessor {
                #(#arms),*
            }
        }
    } else {
        let acc = &accessors[0];
        let accessor_method = format_ident!("{}", acc.parent_method);
        quote! {
            lock.#accessor_method().#method_ident(#(#params_call),*)
        }
    };

    let body: TokenStream = match &m.returns {
        ReturnKind::SelfType => {
            // match block はセミコロン不要、単一式はセミコロン必要
            if multi_accessor {
                quote! {
                    let mut lock = self.parent.lock().unwrap();
                    #accessor_expr
                    #struct_ident { parent: Arc::clone(&self.parent) }
                }
            } else {
                quote! {
                    let mut lock = self.parent.lock().unwrap();
                    #accessor_expr;
                    #struct_ident { parent: Arc::clone(&self.parent) }
                }
            }
        }
        ReturnKind::ResultSelf => {
            quote! {
                let mut lock = self.parent.lock().unwrap();
                #accessor_expr?;
                Ok(#struct_ident { parent: Arc::clone(&self.parent) })
            }
        }
        ReturnKind::ResultVoid => {
            quote! {
                let mut lock = self.parent.lock().unwrap();
                #accessor_expr?;
                Ok(())
            }
        }
        ReturnKind::Void => {
            if multi_accessor {
                quote! {
                    let mut lock = self.parent.lock().unwrap();
                    #accessor_expr
                }
            } else {
                quote! {
                    let mut lock = self.parent.lock().unwrap();
                    #accessor_expr;
                }
            }
        }
        ReturnKind::Other(_) => {
            quote! {
                let mut lock = self.parent.lock().unwrap();
                #accessor_expr
            }
        }
    };

    quote! {
        #[wasm_bindgen(js_name = #js_name, skip_jsdoc)]
        pub fn #method_ident(#sig_params) -> #ret_type {
            #body
        }
    }
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ir::*;

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
        assert!(code.contains("pub struct Format"), "missing pub struct Format in: {}", code);
        assert!(code.contains("pub (crate) inner : Arc < Mutex < xlsx :: Format >>"), "missing inner field in: {}", code);
        assert!(code.contains("derive(Clone)") || code.contains("derive (Clone)"), "missing derive Clone in: {}", code);
        assert!(code.contains("wasm_bindgen"), "missing wasm_bindgen in: {}", code);
    }

    #[test]
    fn constructor_no_params() {
        let s = make_standalone("Format", true);
        let code = generate_struct_file(&s, &CodegenContext::empty());
        assert!(code.contains("wasm_bindgen(constructor)") || code.contains("wasm_bindgen (constructor)"), "missing wasm_bindgen(constructor) in: {}", code);
        assert!(code.contains("pub fn new") && code.contains("Format"), "missing constructor signature in: {}", code);
        assert!(code.contains("Mutex :: new (xlsx :: Format :: new ())") || code.contains("Mutex::new(xlsx::Format::new())"), "missing Mutex::new(xlsx::Format::new()) in: {}", code);
    }

    #[test]
    fn constructor_with_params() {
        let mut s = make_standalone("Note", false);
        s.constructor = Some(AnalyzedConstructor {
            params: vec![AnalyzedParam { name: "text".into(), ty: ParamType::Str }],
        });
        let code = generate_struct_file(&s, &CodegenContext::empty());
        assert!(code.contains("pub fn new") && code.contains("text") && code.contains("& str") && code.contains("Note"), "missing constructor sig in: {}", code);
        assert!(code.contains("xlsx :: Note :: new (text)") || code.contains("xlsx::Note::new(text)"), "missing new call in: {}", code);
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
        assert!(code.contains("pub fn set_bold") && code.contains("& self") && code.contains("Format"), "missing set_bold sig in: {}", code);
        assert!(code.contains("mem :: take") || code.contains("mem::take"), "missing mem::take in: {}", code);
        assert!(code.contains("inner . set_bold") || code.contains("inner.set_bold"), "missing inner.set_bold in: {}", code);
        assert!(code.contains("Arc :: clone (& self . inner)") || code.contains("Arc::clone(&self.inner)"), "missing Arc::clone in: {}", code);
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
        assert!(code.contains("pub fn set_bold") && code.contains("& self") && code.contains("ChartFont"), "missing set_bold sig in: {}", code);
        assert!(code.contains("lock . set_bold") || code.contains("lock.set_bold"), "missing lock.set_bold in: {}", code);
        assert!(!code.contains("mem :: take") && !code.contains("mem::take"), "should not contain mem::take in: {}", code);
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
        assert!(code.contains("pub fn set_font_size") && code.contains("size") && code.contains("f64") && code.contains("Format"), "missing set_font_size sig in: {}", code);
        assert!(code.contains("inner . set_font_size (size)") || code.contains("inner.set_font_size(size)"), "missing inner.set_font_size(size) in: {}", code);
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
        assert!(code.contains("pub fn set_align") && code.contains("align") && code.contains("FormatAlign") && code.contains("Format"), "missing set_align sig in: {}", code);
        assert!(code.contains("xlsx :: FormatAlign :: from (align)") || code.contains("xlsx::FormatAlign::from(align)"), "missing xlsx::FormatAlign::from(align) in: {}", code);
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
        assert!(code.contains("pub fn set_font") && code.contains("font") && code.contains("ChartFont") && code.contains("ChartDataTable"), "missing set_font sig in: {}", code);
        assert!(code.contains("& font . inner") || code.contains("&font.inner"), "missing &font.inner in: {}", code);
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
        assert!(code.contains("WasmResult") && code.contains("Worksheet"), "missing WasmResult<Worksheet> in: {}", code);
        assert!(code.contains("set_column_width (col , width) ?") || code.contains("set_column_width(col, width)?"), "missing set_column_width(col, width)? in: {}", code);
        assert!(code.contains("Ok (Worksheet") || code.contains("Ok(Worksheet"), "missing Ok(Worksheet {{...}}) in: {}", code);
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
        assert!(!code.contains("fn validate"), "should not contain fn validate in: {}", code);
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
        assert!(code.contains("pub (crate) parent : Arc < Mutex < xlsx :: Chart > >") || code.contains("pub(crate) parent: Arc<Mutex<xlsx::Chart>>") || code.contains("pub (crate) parent"), "missing parent field in: {}", code);
        assert!(code.contains("lock . title () . set_name (name)") || code.contains("lock.title().set_name(name)"), "missing lock.title().set_name(name) in: {}", code);
        assert!(code.contains("ChartTitle") && code.contains("Arc :: clone (& self . parent)") || code.contains("Arc::clone(&self.parent)"), "missing ChartTitle with Arc::clone in: {}", code);
    }
}
