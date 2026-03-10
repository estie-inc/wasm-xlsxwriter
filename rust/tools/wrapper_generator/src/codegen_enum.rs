// Generate Rust code for enum wrappers from IR

use proc_macro2::TokenStream;
use quote::{format_ident, quote};

use crate::ir::{AnalyzedEnum, AnalyzedVariant, VariantKind};

pub fn generate_enum_file(e: &AnalyzedEnum) -> String {
    let enum_ident = format_ident!("{}", e.name);

    // Derive macros
    let has_data = e.has_data_variants();
    let core_derives = if has_data {
        quote! { #[derive(Debug, Clone, Serialize, Deserialize, Tsify)] }
    } else {
        quote! { #[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify)] }
    };

    let default_derive = if e.has_default {
        quote! { #[derive(Default)] }
    } else {
        quote! {}
    };

    // Determine which variant gets #[default]
    let default_variant_idx = if e.has_default {
        find_default_variant_idx(e)
    } else {
        None
    };

    // Variants — doc comments are injected as raw /// lines
    let variant_strings: Vec<String> = e
        .variants
        .iter()
        .enumerate()
        .map(|(idx, variant)| {
            let mut parts = Vec::new();
            if let Some(doc) = &variant.doc {
                for line in doc.lines() {
                    parts.push(format!("/// {}", line));
                }
            }
            if default_variant_idx == Some(idx) {
                parts.push("#[default]".to_string());
            }
            let variant_ident = format_ident!("{}", variant.name);
            let variant_tokens = match &variant.kind {
                VariantKind::Plain => {
                    quote! { #variant_ident, }
                }
                VariantKind::Tuple(fields) => {
                    let field_types: Vec<TokenStream> = fields
                        .iter()
                        .map(|ty_str| ty_str.parse::<TokenStream>().expect("valid type token"))
                        .collect();
                    quote! { #variant_ident(#(#field_types),*), }
                }
                VariantKind::Struct(fields) => {
                    let field_tokens: Vec<TokenStream> = fields
                        .iter()
                        .map(|(name, ty_str)| {
                            let field_ident = format_ident!("{}", name);
                            let field_ty: TokenStream =
                                ty_str.parse().expect("valid type token");
                            quote! { #field_ident: #field_ty }
                        })
                        .collect();
                    quote! { #variant_ident { #(#field_tokens),* }, }
                }
            };
            parts.push(variant_tokens.to_string());
            parts.join("\n")
        })
        .collect();
    let variants_body = variant_strings.join("\n");

    // From impl match arms
    let match_arms = e
        .variants
        .iter()
        .map(|variant| generate_match_arm_tokens(&enum_ident, variant))
        .collect::<Vec<_>>();

    let imports = quote! {
        use rust_xlsxwriter as xlsx;
        use serde::{Deserialize, Serialize};
        use tsify::Tsify;
    };

    // Enum doc comment as raw /// lines
    let enum_doc = format_doc_comments(&e.doc);

    let enum_attrs = quote! {
        #core_derives
        #default_derive
        #[tsify(into_wasm_abi, from_wasm_abi)]
    };

    let enum_header = format!(
        "{}{}pub enum {} {{\n{}\n}}",
        enum_doc,
        enum_attrs,
        enum_ident,
        variants_body,
    );

    let from_impl = quote! {
        impl From<#enum_ident> for xlsx::#enum_ident {
            fn from(value: #enum_ident) -> xlsx::#enum_ident {
                match value {
                    #(#match_arms)*
                }
            }
        }
    };

    format!("{}\n\n{}\n\n{}\n", imports, enum_header, from_impl)
}

/// Returns the index of the variant that should get `#[default]`.
/// Prefers a variant named "None" or "Default"; falls back to index 0.
fn find_default_variant_idx(e: &AnalyzedEnum) -> Option<usize> {
    let preferred = e
        .variants
        .iter()
        .position(|v| v.name == "None" || v.name == "Default");
    Some(preferred.unwrap_or(0))
}

fn format_doc_comments(doc: &Option<String>) -> String {
    match doc {
        Some(text) => {
            let mut out = String::new();
            for line in text.lines() {
                out.push_str(&format!("/// {}\n", line));
            }
            out
        }
        None => String::new(),
    }
}

fn generate_match_arm_tokens(
    enum_ident: &proc_macro2::Ident,
    variant: &AnalyzedVariant,
) -> TokenStream {
    let variant_ident = format_ident!("{}", variant.name);
    match &variant.kind {
        VariantKind::Plain => {
            quote! {
                #enum_ident::#variant_ident => xlsx::#enum_ident::#variant_ident,
            }
        }
        VariantKind::Tuple(fields) => {
            let bindings: Vec<proc_macro2::Ident> =
                (0..fields.len()).map(|i| format_ident!("v{}", i)).collect();
            quote! {
                #enum_ident::#variant_ident(#(#bindings),*) => xlsx::#enum_ident::#variant_ident(#(#bindings),*),
            }
        }
        VariantKind::Struct(fields) => {
            let field_idents: Vec<proc_macro2::Ident> = fields
                .iter()
                .map(|(name, _)| format_ident!("{}", name))
                .collect();
            quote! {
                #enum_ident::#variant_ident { #(#field_idents),* } => xlsx::#enum_ident::#variant_ident { #(#field_idents),* },
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ir::*;

    #[test]
    fn plain_enum_generates_tsify() {
        let e = AnalyzedEnum {
            name: "FormatAlign".into(),
            variants: vec![
                AnalyzedVariant { name: "General".into(), kind: VariantKind::Plain, doc: None },
                AnalyzedVariant { name: "Left".into(), kind: VariantKind::Plain, doc: None },
                AnalyzedVariant { name: "Center".into(), kind: VariantKind::Plain, doc: None },
            ],
            has_default: false,
            doc: None,
        };
        let code = generate_enum_file(&e);
        assert!(code.contains("derive (Debug , Clone , Copy , Serialize , Deserialize , Tsify)"));
        assert!(code.contains("tsify (into_wasm_abi , from_wasm_abi)"));
        assert!(code.contains("pub enum FormatAlign"));
        assert!(code.contains("General ,"));
        assert!(code.contains("Left ,"));
        assert!(code.contains("Center ,"));
        // NOT wasm_bindgen
        assert!(!code.contains("wasm_bindgen"));
    }

    #[test]
    fn enum_from_impl_plain() {
        let e = AnalyzedEnum {
            name: "FormatBorder".into(),
            variants: vec![
                AnalyzedVariant { name: "None".into(), kind: VariantKind::Plain, doc: None },
                AnalyzedVariant { name: "Thin".into(), kind: VariantKind::Plain, doc: None },
            ],
            has_default: false,
            doc: None,
        };
        let code = generate_enum_file(&e);
        assert!(code.contains("impl From < FormatBorder > for xlsx :: FormatBorder"));
        assert!(code.contains("FormatBorder :: None => xlsx :: FormatBorder :: None"));
        assert!(code.contains("FormatBorder :: Thin => xlsx :: FormatBorder :: Thin"));
    }

    #[test]
    fn enum_with_data_variants() {
        let e = AnalyzedEnum {
            name: "Color".into(),
            variants: vec![
                AnalyzedVariant { name: "Default".into(), kind: VariantKind::Plain, doc: None },
                AnalyzedVariant { name: "Red".into(), kind: VariantKind::Plain, doc: None },
                AnalyzedVariant { name: "RGB".into(), kind: VariantKind::Tuple(vec!["u32".into()]), doc: None },
                AnalyzedVariant { name: "Theme".into(), kind: VariantKind::Tuple(vec!["u8".into(), "u8".into()]), doc: None },
            ],
            has_default: false,
            doc: None,
        };
        let code = generate_enum_file(&e);
        // No Copy derive (has data variants)
        assert!(!code.contains("Copy"));
        assert!(code.contains("RGB (u32)"));
        assert!(code.contains("Theme (u8 , u8)"));
        // From impl handles data
        assert!(code.contains("Color :: RGB (v0) => xlsx :: Color :: RGB (v0)"));
        assert!(code.contains("Color :: Theme (v0 , v1) => xlsx :: Color :: Theme (v0 , v1)"));
    }

    #[test]
    fn enum_with_default() {
        let e = AnalyzedEnum {
            name: "Color".into(),
            variants: vec![
                AnalyzedVariant { name: "Default".into(), kind: VariantKind::Plain, doc: None },
                AnalyzedVariant { name: "Red".into(), kind: VariantKind::Plain, doc: None },
            ],
            has_default: true,
            doc: None,
        };
        let code = generate_enum_file(&e);
        assert!(code.contains("derive (Default)"));
        assert!(code.contains("#[default]"));
    }

    #[test]
    fn enum_with_doc_comments() {
        let e = AnalyzedEnum {
            name: "ChartType".into(),
            variants: vec![
                AnalyzedVariant { name: "Area".into(), kind: VariantKind::Plain, doc: Some("Area chart.".into()) },
                AnalyzedVariant { name: "Bar".into(), kind: VariantKind::Plain, doc: Some("Bar chart.".into()) },
            ],
            has_default: false,
            doc: Some("The type of chart.".into()),
        };
        let code = generate_enum_file(&e);
        assert!(code.contains("/// The type of chart."));
        assert!(code.contains("/// Area chart."));
        assert!(code.contains("/// Bar chart."));
    }

    #[test]
    fn enum_use_statements() {
        let e = AnalyzedEnum {
            name: "FormatAlign".into(),
            variants: vec![
                AnalyzedVariant { name: "Left".into(), kind: VariantKind::Plain, doc: None },
            ],
            has_default: false,
            doc: None,
        };
        let code = generate_enum_file(&e);
        assert!(code.contains("use rust_xlsxwriter as xlsx"));
        assert!(code.contains("use serde"));
        assert!(code.contains("Deserialize"));
        assert!(code.contains("Serialize"));
        assert!(code.contains("use tsify :: Tsify"));
    }
}
