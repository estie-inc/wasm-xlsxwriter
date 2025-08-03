use crate_inspector::{CrateItem, EnumItem};
use ruast::*;

use crate::utils::{new_line, process_doc_comment};

pub fn generate_enum_wrapper_output( enum_info: &EnumItem) -> Crate {
    let mut krate = Crate::new();
    let uses = crate::common::generate_enum_use_statements();
    let enum_ = generate_enum_wrapper(&enum_info);
    let from_impl = generate_enum_impl_from(enum_info);

    // Build crate
    uses.into_iter().for_each(|use_item| {
        krate.add_item(use_item);
    });
    krate.add_item(enum_);
    krate.add_item(from_impl);
    
    krate
}

fn generate_enum_wrapper(enum_info: &EnumItem) -> Item<ItemKind> {
    let mut enum_def = EnumDef::empty(enum_info.name());
    
    // Add variants from original enum with their documentation
    for variant in enum_info.variants() {
        let variant_name = variant.name();
        let variant_def = Variant::empty(variant_name);
        
        // TODO: Add documentation for variants when crate_inspector supports it
        // Currently VariantItem doesn't expose docs field
        
        enum_def.add_variant(variant_def);
    }

    let mut item = Item::from(enum_def);
    item.vis = Visibility::Public;
    item.add_attr(new_line());
    
    // Add documentation if available
    if let Some(doc) = &enum_info.item().docs {
        item.add_attr(Attribute::doc_comment(process_doc_comment(doc)));
    }
    
    // Add derive attributes
    item.add_attr(Attribute::normal(AttributeItem::new(
        "derive",
        AttrArgs::Delimited(DelimArgs::parenthesis(
            vec![
                Token::Ident("Debug".to_string()).into_joint(),
                Token::Comma,
                Token::Ident("Clone".to_string()).into_joint(),
                Token::Comma,
                Token::Ident("Copy".to_string()).into_joint(),
                Token::Comma,
                Token::Ident("Default".to_string()),
            ]
            .into_tokens(),
        )),
    )));
    
    // Add wasm_bindgen attribute
    item.add_attr(Attribute::normal(AttributeItem::new(
        "wasm_bindgen",
        AttrArgs::Empty,
    )));

    item
}

fn generate_enum_impl_from(enum_info: &EnumItem) -> Item<ItemKind> {
    let enum_name = enum_info.name();
    
    // Generate match arms for each variant
    let mut variants_match_arms: Vec<Arm> = vec![];
    
    for variant in enum_info.variants() {
        let variant_name = variant.name();
        
        // Create pattern: EnumName::VariantName
        let pattern = Pat::ident(format!("{}::{}", enum_name, variant_name));
        
        // Create expression: xlsx::EnumName::VariantName
        let expression = Expr::from(ExprKind::Path(
            Path::single("xlsx").chain(enum_name).chain(variant_name)
        ));
        
        variants_match_arms.push(Arm::new(pattern, None, expression));
    }
    
    let match_expr = Match::new(
        Expr::from(ExprKind::Path(Path::single("value"))),
        variants_match_arms
    );
    
    let from_impl = Impl::new(
        vec![],
        Some(Type::Path(Path::single(PathSegment::new("From", Some(vec![GenericArg::Type(
            Type::Path(Path::single(PathSegment::new(enum_name, None))),
        )]))))),
        Type::Path(Path::single(PathSegment::new("xlsx", None)).chain(PathSegment::new(enum_name, None))),
        None,
        vec![AssocItem::from(AssocItemKind::Fn(Fn::simple(
            "from",
            FnDecl::regular(
                vec![Param::new(Pat::ident("value"), Type::Path(Path::single(enum_name)))],
                Some(Type::Path(Path::single("Self")))
            ),
            Block::single(ExprKind::Match(match_expr)),
        )))],
    );

    Item::from(from_impl)
}