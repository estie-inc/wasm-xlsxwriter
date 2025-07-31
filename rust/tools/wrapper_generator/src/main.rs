use anyhow::{Context, Result};
use clap::Parser;
use crate_inspector::{CrateBuilder, CrateItem, StructItem};
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

    let krate = get_crate_info(&args.manifest_path, false).context("Failed to get crate info")?;
    let struct_info = krate
        .all_structs()
        .find(|s| s.name() == args.struct_name)
        .context(format!("Struct '{}' not found", args.struct_name))?;

    let mut krate = Crate::new();

    let uses = generate_use_statements();
    let struct_ = generate_struct_wrapper(&struct_info);

    let common_methods = generate_common_methods(&struct_info);
    let wrapper_methods = generate_wrapper_methods(&struct_info);

    let from_impls = generate_impl_from(struct_info.name());

    let mut impl_block = Item::from(Impl::simple(
        Type::Path(Path::single(PathSegment::new(struct_info.name(), None))),
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
    from_impls.into_iter().for_each(|from_impl| {
        krate.add_item(from_impl);
    });
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

fn generate_struct_wrapper(struct_info: &StructItem) -> Item<ItemKind> {
    let mut struct_def = StructDef::empty(struct_info.name());
    struct_def.add_field(FieldDef::new(
        Visibility::crate_(),
        Some("inner"),
        Type::poly_path(
            "Arc",
            vec![GenericArg::Type(Type::poly_path(
                "Mutex",
                vec![GenericArg::Type(Type::Path(
                    Path::single("xlsx").chain(struct_info.name()),
                ))],
            ))],
        ),
    ));

    let mut item = Item::from(struct_def);
    item.vis = Visibility::Public;
    item.add_attr(new_line());
    if let Some(doc) = &struct_info.item().docs {
        item.add_attr(Attribute::doc_comment(process_doc_comment(doc)));
    }
    item.add_attr(Attribute::normal(AttributeItem::new(
        "derive",
        AttrArgs::Delimited(DelimArgs::parenthesis(
            vec![
                Token::Ident("Debug".to_string()).into_joint(),
                Token::Comma,
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

fn generate_impl_from(struct_name: &str) -> Vec<Item<ItemKind>> {
    // First implementation: impl From<StructName> for xlsx::StructName
    let from_impl1 = Impl::new(
        vec![],
        Some(Type::Path(Path::single(PathSegment::new("From", Some(vec![GenericArg::Type(
            Type::Path(Path::single(PathSegment::new(struct_name, None))),
        )]))))),
        Type::Path(Path::single(PathSegment::new("xlsx", None)).chain(PathSegment::new(struct_name, None))),
        None,
        vec![AssocItem::from(AssocItemKind::Fn(Fn::simple(
            "from",
            FnDecl::regular(
                vec![Param::new(Pat::ident("format"), Type::Path(Path::single(struct_name)))],
                Some(Type::Path(Path::single("Self")))
            ),
            Block::single(
                ExprKind::method_call0(
                    ExprKind::from(ExprKind::method_call0(
                        ExprKind::Path(Path::single("format")),
                        "lock"
                    )),
                    "clone"
                )
            ),
        )))],
    );

    // Second implementation: impl From<StructName> for &'static xlsx::StructName
    let from_impl2 = Impl::new(
        vec![],
        Some(Type::Path(Path::single(PathSegment::new("From", Some(vec![GenericArg::Type(
            Type::Path(Path::single(PathSegment::new(struct_name, None))),
        )]))))),
        Type::static_ref(Type::Path(Path::single(PathSegment::new("xlsx", None)).chain(PathSegment::new(struct_name, None)))),
        None,
        vec![AssocItem::from(AssocItemKind::Fn(Fn::simple(
            "from",
            FnDecl::regular(
                vec![Param::new(Pat::ident("format"), Type::Path(Path::single(struct_name)))],
                Some(Type::Path(Path::single("Self")))
            ),
            Block::single(
                ExprKind::call(
                    ExprKind::Path(Path::single("Box").chain("leak")),
                    vec![
                        Expr::from(ExprKind::call(
                            ExprKind::Path(Path::single("Box").chain("new")),
                            vec![
                                Expr::from(ExprKind::method_call0(
                                    ExprKind::from(ExprKind::method_call0(
                                        ExprKind::Path(Path::single("format")),
                                        "lock"
                                    )),
                                    "clone"
                                ))
                            ]
                        ))
                    ]
                )
            ),
        )))],
    );

    // Return both implementations as a vector
    vec![Item::from(from_impl1), Item::from(from_impl2)]
}

fn get_crate_info(manifest_path: &str, document_private: bool) -> Result<crate_inspector::Crate> {
    Ok(CrateBuilder::default()
        .manifest_path(manifest_path)
        .document_private_items(document_private)
        .build()
        .context("Failed to build crate inspector")?)
}
