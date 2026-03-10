// IR から enum ラッパーの Rust コードを生成する

use crate::ir::{AnalyzedEnum, VariantKind};

pub fn generate_enum_file(e: &AnalyzedEnum) -> String {
    let mut out = String::new();

    // Imports
    out.push_str("use rust_xlsxwriter as xlsx;\n");
    out.push_str("use serde::{Deserialize, Serialize};\n");
    out.push_str("use tsify::Tsify;\n");
    out.push('\n');

    // Enum-level doc comment
    if let Some(doc) = &e.doc {
        for line in doc.lines() {
            out.push_str(&format!("/// {}\n", line));
        }
    }

    // Derive macros
    let has_data = e.has_data_variants();
    let mut derives = vec!["Debug", "Clone"];
    if !has_data {
        derives.push("Copy");
    }
    derives.extend(["Serialize", "Deserialize", "Tsify"]);
    out.push_str(&format!("#[derive({})]\n", derives.join(", ")));
    if e.has_default {
        out.push_str("#[derive(Default)]\n");
    }
    out.push_str("#[tsify(into_wasm_abi, from_wasm_abi)]\n");
    out.push_str(&format!("pub enum {} {{\n", e.name));

    // Determine which variant gets #[default]
    let default_variant_idx = if e.has_default {
        find_default_variant_idx(e)
    } else {
        None
    };

    // Variants
    for (idx, variant) in e.variants.iter().enumerate() {
        if let Some(doc) = &variant.doc {
            for line in doc.lines() {
                out.push_str(&format!("    /// {}\n", line));
            }
        }
        if default_variant_idx == Some(idx) {
            out.push_str("    #[default]\n");
        }
        match &variant.kind {
            VariantKind::Plain => {
                out.push_str(&format!("    {},\n", variant.name));
            }
            VariantKind::Tuple(fields) => {
                let fields_str = fields.join(", ");
                out.push_str(&format!("    {}({}),\n", variant.name, fields_str));
            }
            VariantKind::Struct(fields) => {
                let fields_str = fields
                    .iter()
                    .map(|(name, ty)| format!("{}: {}", name, ty))
                    .collect::<Vec<_>>()
                    .join(", ");
                out.push_str(&format!("    {} {{ {} }},\n", variant.name, fields_str));
            }
        }
    }

    out.push_str("}\n");
    out.push('\n');

    // From impl
    out.push_str(&format!(
        "impl From<{name}> for xlsx::{name} {{\n",
        name = e.name
    ));
    out.push_str(&format!(
        "    fn from(value: {name}) -> xlsx::{name} {{\n",
        name = e.name
    ));
    out.push_str("        match value {\n");

    for variant in &e.variants {
        let arm = generate_match_arm(&e.name, variant);
        out.push_str(&format!("            {}\n", arm));
    }

    out.push_str("        }\n");
    out.push_str("    }\n");
    out.push_str("}\n");

    out
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

fn generate_match_arm(enum_name: &str, variant: &crate::ir::AnalyzedVariant) -> String {
    match &variant.kind {
        VariantKind::Plain => {
            format!(
                "{name}::{var} => xlsx::{name}::{var},",
                name = enum_name,
                var = variant.name
            )
        }
        VariantKind::Tuple(fields) => {
            let bindings: Vec<String> = (0..fields.len()).map(|i| format!("v{}", i)).collect();
            let bindings_str = bindings.join(", ");
            format!(
                "{name}::{var}({binds}) => xlsx::{name}::{var}({binds}),",
                name = enum_name,
                var = variant.name,
                binds = bindings_str
            )
        }
        VariantKind::Struct(fields) => {
            let field_names: Vec<String> = fields.iter().map(|(name, _)| name.clone()).collect();
            let fields_str = field_names.join(", ");
            format!(
                "{name}::{var} {{ {fields} }} => xlsx::{name}::{var} {{ {fields} }},",
                name = enum_name,
                var = variant.name,
                fields = fields_str
            )
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
        assert!(code.contains("#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify)]"));
        assert!(code.contains("#[tsify(into_wasm_abi, from_wasm_abi)]"));
        assert!(code.contains("pub enum FormatAlign {"));
        assert!(code.contains("General,"));
        assert!(code.contains("Left,"));
        assert!(code.contains("Center,"));
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
        assert!(code.contains("impl From<FormatBorder> for xlsx::FormatBorder {"));
        assert!(code.contains("FormatBorder::None => xlsx::FormatBorder::None,"));
        assert!(code.contains("FormatBorder::Thin => xlsx::FormatBorder::Thin,"));
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
        assert!(code.contains("RGB(u32),"));
        assert!(code.contains("Theme(u8, u8),"));
        // From impl handles data
        assert!(code.contains("Color::RGB(v0) => xlsx::Color::RGB(v0),"));
        assert!(code.contains("Color::Theme(v0, v1) => xlsx::Color::Theme(v0, v1),"));
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
        assert!(code.contains("Default)]")); // derive(Default)
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
        assert!(code.contains("use rust_xlsxwriter as xlsx;"));
        assert!(code.contains("use serde::{Deserialize, Serialize};"));
        assert!(code.contains("use tsify::Tsify;"));
    }
}
