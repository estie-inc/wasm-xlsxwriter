/// quote! ベースのコード生成で共有するトークンヘルパー。
///
/// ParamType → TokenStream の変換を一元化し、codegen / codegen_enum の両方から利用する。
use proc_macro2::TokenStream;
use quote::{format_ident, quote};

use crate::codegen::CodegenContext;
use crate::ir::ParamType;

/// ParamType を wasm-bindgen シグネチャ用の型トークンに変換
pub fn param_type_tokens(ty: &ParamType) -> TokenStream {
    match ty {
        ParamType::Bool => quote! { bool },
        ParamType::U8 => quote! { u8 },
        ParamType::U16 => quote! { u16 },
        ParamType::U32 => quote! { u32 },
        ParamType::U64 => quote! { u64 },
        ParamType::I8 => quote! { i8 },
        ParamType::I16 => quote! { i16 },
        ParamType::I32 => quote! { i32 },
        ParamType::I64 => quote! { i64 },
        ParamType::F32 => quote! { f32 },
        ParamType::F64 => quote! { f64 },
        ParamType::Usize => quote! { usize },
        ParamType::Str => quote! { &str },
        ParamType::WrappedType(name) | ParamType::RefWrappedType(name) => {
            let ident = format_ident!("{}", name);
            quote! { #ident }
        }
        ParamType::VecOf(inner) | ParamType::RefSliceOf(inner) => {
            let inner_tokens = param_type_tokens(inner);
            quote! { Vec<#inner_tokens> }
        }
        ParamType::OptionOf(inner) => {
            let inner_tokens = param_type_tokens(inner);
            quote! { Option<#inner_tokens> }
        }
        ParamType::Unknown(s) => {
            let tokens: TokenStream = s.parse().expect("Unknown type should be valid tokens");
            quote! { #tokens }
        }
    }
}

/// パラメータのシグネチャトークン (e.g., `size: f64`)
pub fn param_sig_tokens(name: &str, ty: &ParamType) -> TokenStream {
    let name_ident = format_ident!("{}", name);
    let ty_tokens = param_type_tokens(ty);
    quote! { #name_ident: #ty_tokens }
}

/// パラメータの呼び出し式トークン (enum → xlsx::T::from, struct → inner access, etc.)
pub fn param_call_tokens(name: &str, ty: &ParamType, ctx: &CodegenContext) -> TokenStream {
    let name_ident = format_ident!("{}", name);
    match ty {
        ParamType::WrappedType(type_name) => {
            if ctx.handwritten_enum_names.contains(type_name) {
                quote! { #name_ident.inner }
            } else if ctx.enum_names.contains(type_name) {
                let xlsx_ty = format_ident!("{}", type_name);
                quote! { xlsx::#xlsx_ty::from(#name_ident) }
            } else {
                struct_inner_owned_tokens(&name_ident, type_name, ctx)
            }
        }
        ParamType::RefWrappedType(type_name) => {
            if ctx.handwritten_enum_names.contains(type_name) {
                quote! { &#name_ident.inner }
            } else if ctx.enum_names.contains(type_name) {
                let xlsx_ty = format_ident!("{}", type_name);
                quote! { &xlsx::#xlsx_ty::from(#name_ident) }
            } else {
                struct_inner_ref_tokens(&name_ident, type_name, ctx)
            }
        }
        ParamType::VecOf(inner) => match inner.as_ref() {
            ParamType::WrappedType(type_name) => {
                vec_of_wrapped_tokens(&name_ident, type_name, ctx)
            }
            _ => quote! { #name_ident },
        },
        ParamType::RefSliceOf(inner) => match inner.as_ref() {
            ParamType::WrappedType(type_name) => {
                ref_slice_of_wrapped_tokens(&name_ident, type_name, ctx)
            }
            _ => quote! { &#name_ident },
        },
        _ => quote! { #name_ident },
    }
}

// ── private helpers ──────────────────────────────────────────────────

fn struct_inner_owned_tokens(
    name_ident: &proc_macro2::Ident,
    type_name: &str,
    ctx: &CodegenContext,
) -> TokenStream {
    if ctx.handwritten_struct_names.contains(type_name) {
        quote! { #name_ident.inner.clone() }
    } else {
        quote! { #name_ident.inner.lock().unwrap().clone() }
    }
}

fn struct_inner_ref_tokens(
    name_ident: &proc_macro2::Ident,
    type_name: &str,
    ctx: &CodegenContext,
) -> TokenStream {
    if ctx.handwritten_struct_names.contains(type_name) {
        quote! { &#name_ident.inner }
    } else {
        quote! { &*#name_ident.inner.lock().unwrap() }
    }
}

fn vec_of_wrapped_tokens(
    name_ident: &proc_macro2::Ident,
    type_name: &str,
    ctx: &CodegenContext,
) -> TokenStream {
    if ctx.handwritten_enum_names.contains(type_name) {
        quote! { #name_ident.into_iter().map(|x| x.inner).collect() }
    } else if ctx.enum_names.contains(type_name) {
        let xlsx_ty = format_ident!("{}", type_name);
        quote! { #name_ident.into_iter().map(|x| xlsx::#xlsx_ty::from(x)).collect() }
    } else if ctx.handwritten_struct_names.contains(type_name) {
        quote! { #name_ident.iter().map(|x| x.inner.clone()).collect() }
    } else {
        quote! { #name_ident.iter().map(|x| x.inner.lock().unwrap().clone()).collect() }
    }
}

fn ref_slice_of_wrapped_tokens(
    name_ident: &proc_macro2::Ident,
    type_name: &str,
    ctx: &CodegenContext,
) -> TokenStream {
    if ctx.handwritten_enum_names.contains(type_name) {
        quote! { &#name_ident.iter().map(|x| x.inner).collect::<Vec<_>>() }
    } else if ctx.enum_names.contains(type_name) {
        let xlsx_ty = format_ident!("{}", type_name);
        quote! { &#name_ident.into_iter().map(|x| xlsx::#xlsx_ty::from(x)).collect::<Vec<_>>() }
    } else if ctx.handwritten_struct_names.contains(type_name) {
        quote! { &#name_ident.iter().map(|x| x.inner.clone()).collect::<Vec<_>>() }
    } else {
        quote! { &#name_ident.iter().map(|x| x.inner.lock().unwrap().clone()).collect::<Vec<_>>() }
    }
}
