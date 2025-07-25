use anyhow::{Context, Result};
use clap::Parser;
use crate_info_extractor::{extract_crate_items, get_crate_info, StructInfo};
use method_generator::{generate_common_methods, generate_wrapper_methods};
use ruast::*;
use std::path::PathBuf;

mod method_generator;
mod utils;

use crate::utils::{new_line, process_doc_comment};

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

    let crate_info = get_crate_info(&args.manifest_path).context("Failed to get crate info")?;
    let extracted_items = extract_crate_items(&crate_info);
    let struct_info = extracted_items
        .structs
        .iter()
        .find(|s| s.name == args.struct_name)
        .context(format!("Struct '{}' not found", args.struct_name))?;

    let mut krate = Crate::new();

    let uses = generate_use_statements();
    let struct_ = generate_struct_wrapper(struct_info);

    let common_methods = generate_common_methods(&struct_info);
    let wrapper_methods = generate_wrapper_methods(&struct_info);

    let mut impl_block = Item::from(Impl::simple(
        Type::Path(Path::single(PathSegment::new(&struct_info.name, None))),
        common_methods.into_iter().chain(wrapper_methods).collect(),
    ));
    impl_block.add_attr(new_line());
    impl_block.add_attr(Attribute::normal(AttributeItem::new(
        "wasm_bindgen",
        AttrArgs::Empty,
    )));

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
        // use rust_xlsxwriter as xlsx;
        Use::from(Path::single("rust_xlsxwriter").chain_use_rename("xlsx")),
        // use std::sync::{Arc, Mutex, MutexGuard}
        Use::from(Path::single("std").chain("sync").chain_use_group(vec![
            UseTree::from(Path::single(PathSegment::new("Arc", None))),
            UseTree::from(Path::single(PathSegment::new("Mutex", None))),
            UseTree::from(Path::single(PathSegment::new("MutexGuard", None))),
        ])),
        // use wasm_bindgen::prelude::*;
        Use::from(
            Path::single("wasm_bindgen")
                .chain("prelude")
                .chain_use_glob(),
        ),
        // use crate::impl_method;
        Use::from(Path::single("crate").chain("impl_method")),
    ]
}

fn generate_struct_wrapper(struct_info: &StructInfo) -> Item<ItemKind> {
    let mut struct_def = StructDef::empty(&struct_info.name);
    struct_def.add_field(FieldDef::new(
        Visibility::crate_(),
        Some("inner"),
        Type::poly_path(
            "Arc",
            vec![GenericArg::Type(Type::poly_path(
                "Mutex",
                vec![GenericArg::Type(Type::Path(
                    Path::single("xlsx").chain(&struct_info.name),
                ))],
            ))],
        ),
    ));

    let mut item = Item::from(struct_def);
    item.add_attr(new_line());
    if let Some(doc) = &struct_info.doc {
        item.add_attr(Attribute::doc_comment(process_doc_comment(doc)));
    }
    item.add_attr(Attribute::normal(AttributeItem::new(
        "derive",
        AttrArgs::Delimited(DelimArgs::parenthesis(
            vec![
                Token::Ident("Debug".to_string()),
                Token::Comma,
                Token::DocComment(" ".to_string()),
                Token::Ident("Clone".to_string()),
            ]
            .into_tokens(),
        )),
    )));
    item.add_attr(Attribute::normal(AttributeItem::new(
        "wasm_bindgen",
        AttrArgs::Empty,
    )));

    item
}
