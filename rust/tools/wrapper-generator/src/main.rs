use anyhow::{Context, Result};
use clap::Parser;
use crate_info_extractor::{extract_crate_items, get_crate_info};
use ruast::*;
use std::path::PathBuf;

/// Tool to automatically generate wasm_xlsxwriter wrapper methods from rust_xlsxwriter
#[derive(Parser, Debug)]
#[command(author, version, about, long_about = None)]
struct Args {
    /// Path to the Cargo.toml of the crate to extract methods from
    #[arg(short, long)]
    manifest_path: String,

    /// Name of the struct to generate wrappers for
    #[arg(short, long)]
    struct_name: String,

    /// Output file path (optional, defaults to stdout)
    #[arg(short, long)]
    output: Option<PathBuf>,
}

fn main() -> Result<()> {
    let args = Args::parse();

    // Get crate info
    let crate_info = get_crate_info(&args.manifest_path).context("Failed to get crate info")?;

    // Extract items from the crate
    let extracted_items = extract_crate_items(&crate_info);

    // Find the specified struct
    let struct_info = extracted_items
        .structs
        .iter()
        .find(|s| s.name == args.struct_name)
        .context(format!("Struct '{}' not found", args.struct_name))?;

    // Generate wrapper methods
    let mut krate = Crate::new();
    let struct_ = generate_struct_wrapper(&args.struct_name);

    krate.add_item(struct_);

    println!("{}", krate);

    Ok(())
}

fn generate_struct_wrapper(struct_name: &str) -> Item<ItemKind> {
    let mut struct_def = StructDef::empty(struct_name.to_string());
    struct_def.add_field(FieldDef::new(
        Visibility::Scoped(Path::single(PathSegment::new("crate", None))),
        Some("inner"),
        Type::poly_path(
            "Arc",
            vec![GenericArg::Type(Type::poly_path(
                "Mutex",
                vec![GenericArg::Type(Type::Path(Path::new(vec![
                    PathSegment::new("xlsx", None),
                    PathSegment::new(struct_name, None),
                ])))],
            ))],
        ),
    ));

    let mut item = Item::from(struct_def);
    item.add_attr(Attribute::new(AttrKind::Normal(AttributeItem::new(
        "Derive",
        AttrArgs::Delimited(DelimArgs::parenthesis(
            vec![
                Token::Ident("Debug".to_string()),
                Token::Comma,
                Token::Ident("Clone".to_string()),
            ]
            .into_tokens(),
        )),
    ))));
    item.add_attr(Attribute::new(AttrKind::Normal(AttributeItem::new(
        "wasm_bindgen",
        AttrArgs::Empty,
    ))));

    item
}

// fn impl_common_methods(struct_name: &str) -> Item<ItemKind> {
//     let methods = vec![
//         MethodInfo {
//             name: "new".to_string(),
//             args: vec![],
//             return_type: Some(format!("{}::new", struct_name)),
//             body: "Self { inner: xlsx::{}::new() }".to_string(),
//         },
//         MethodInfo {
//             name: "default".to_string(),
//             args: vec![],
//             return_type: Some(format!("{}::default", struct_name)),
//             body: "Self { inner: xlsx::{}::default() }".to_string(),
//         },
//     ];
//
//     let mut impl_block = ImplBlock::new(struct_name.to_string());
//     for method in methods {
//         impl_block.add_method(Method::new(
//             method.name,
//             method.args,
//             method.return_type,
//             method.body,
//         ));
//     }
//
//     Item::from(impl_block)
// }

/// Convert a snake_case string to camelCase
fn to_camel_case(s: &str) -> String {
    let mut result = String::new();
    let mut capitalize_next = false;

    for c in s.chars() {
        if c == '_' {
            capitalize_next = true;
        } else if capitalize_next {
            result.push(c.to_ascii_uppercase());
            capitalize_next = false;
        } else {
            result.push(c);
        }
    }

    result
}
