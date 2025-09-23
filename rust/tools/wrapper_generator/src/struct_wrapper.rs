use crate_inspector::{CrateItem, StructItem};
use crate::method_wrapper::{generate_common_methods, generate_wrapper_methods};
use ruast::*;

use crate::utils::{new_line, process_doc_comment};

pub fn generate_struct_wrapper_output(struct_info: &StructItem) -> Crate {
    let mut krate = Crate::new();
    let uses = crate::common::generate_use_statements();
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
    
    return krate
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