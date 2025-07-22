use anyhow::{Context, Result};
use clap::Parser;
use crate_info_extractor::{extract_crate_items, get_crate_info, StructInfo};
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

    let mut krate = Crate::new();

    let uses = generate_use_statements();
    let struct_ = generate_struct_wrapper(&struct_info.name);
    let impl_block = impl_common_methods(&struct_info);

    // Build crate
    uses.into_iter().for_each(|use_item| {
        krate.add_item(use_item);
    });
    krate.add_item(struct_);
    krate.add_item(impl_block);

    println!("{}", krate);

    Ok(())
}

fn generate_use_statements() -> Vec<Use> {
    vec![
        // use std::sync::{Arc, Mutex, MutexGuard}
        Use::path(Path::new(vec![
            PathSegment::new("std", None),
            PathSegment::new("sync", None),
            PathSegment::new(
                UseTree::group(vec![
                    UseTree::path(Path::new(vec![PathSegment::new("Arc", None)])),
                    UseTree::path(Path::new(vec![PathSegment::new("Mutex", None)])),
                    UseTree::path(Path::new(vec![PathSegment::new("MutexGuard", None)])),
                ])
                .to_string(), // FIXME: UseTree is converted to string here
                None,
            ),
        ])),
        // use wasm_bindgen::prelude::*;
        Use::path(Path::new(vec![
            PathSegment::new("wasm_bindgen", None),
            PathSegment::new("prelude", None),
            PathSegment::new("*", None),
        ])),
        // use rust_xlsxwriter as xlsx;
        Use::rename(UseRename {
            path: Path::new(vec![PathSegment::new("rust_xlsxwriter", None)]),
            alias: "xlsx".to_string(),
        }),
    ]
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

/// Generate common methods for the struct
/// lock, deep_clone
fn impl_common_methods(struct_info: &StructInfo) -> Item<ItemKind> {
    let mut lock_fn = Item::from(Fn::simple(
        "lock".to_string(),
        FnDecl::regular(
            vec![Param::ref_self()],
            Some(Type::Path(Path::new(vec![PathSegment::new(
                "MutexGuard",
                None,
            )]))),
        ),
        Block::empty(),
        // new(vec![Stmt::Expr(Expr::MethodCall(MethodCall {
        //             receiver: Expr::Field(FieldAccess {
        //                 expr: Expr::SelfValue,
        //                 field: "inner".to_string(),
        //             }),
        //             method: "lock".to_string(),
        //             args: vec![],
        //         }))])
    ));
    lock_fn.vis = Visibility::Scoped(Path::single(PathSegment::new("crate", None)));

    let impl_block = Impl::simple(
        Type::Path(Path::single(PathSegment::new(&struct_info.name, None))), // The struct name
        vec![AssocItem::from(lock_fn)],
    );
    let mut item = Item::from(impl_block);
    item.add_attr(Attribute::new(AttrKind::Normal(AttributeItem::new(
        "wasm_bindgen",
        AttrArgs::Empty,
    ))));
    item
}

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
