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
        Use::from(Path::single("std").chain("sync").chain_use_group(vec![
                UseTree::from(Path::single(PathSegment::new("Arc", None))),
                UseTree::from(Path::single(PathSegment::new("Mutex", None))),
                UseTree::from(Path::single(PathSegment::new("MutexGuard", None))),
        ])),
        // use wasm_bindgen::prelude::*;
        Use::from(Path::single("wasm_bindgen").chain("prelude").chain_use_glob()),
        // use rust_xlsxwriter as xlsx;
        Use::from(Path::single("rust_xlsxwriter").chain("xlsx").chain_use_rename("xlsx")),
    ]
}

fn generate_struct_wrapper(struct_name: &str) -> Item<ItemKind> {
    let mut struct_def = StructDef::empty(struct_name.to_string());
    struct_def.add_field(FieldDef::new(
        Visibility::crate_(),
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
        Block::single(Expr::new(ExprKind::MethodCall(MethodCall::new(
            Expr::new(ExprKind::MethodCall(MethodCall::new(
                Expr::new(ExprKind::Field(Field::new(
                    Expr::new(ExprKind::Path(
                        Path::single(PathSegment::new("self", None))
                    )),
                    "inner".to_string(),
                ))),
                "lock".to_string(),
                vec![],
            ))),
            "unwrap".to_string(),
            vec![],
        )))),
    ));
    lock_fn.vis = Visibility::crate_();

    // let lock_fn2 = Item::from(Fn::simple(
    //     "lock".to_string(),
    //     FnDecl::regular(
    //         vec![Param::ref_self()],
    //         Some(Type::Path(Path::new(vec![PathSegment::new(
    //             "String",
    //             None,
    //         )]))),
    //     ),
    //     Block::single(Expr::new(ExprKind::MethodCall(MethodCall::new(
    //         Expr::new(ExprKind::MethodCall(MethodCall::new(
    //             Expr::new(Expr::add(
    //                 Expr::new(Lit::int("1".to_string())),
    //                 Expr::new(Lit::int("2".to_string())),
    //             )),
    //             "lock".to_string(),
    //             vec![],
    //         ))),
    //         "unwrap".to_string(),
    //         vec![],
    //     )))),
    // ));


    let impl_block = Impl::simple(
        Type::Path(Path::single(PathSegment::new(&struct_info.name, None))), // The struct name
        vec![lock_fn],
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
