use crate::utils::{new_line, process_doc_comment, to_camel_case};
use crate_inspector::{CrateItem, FunctionItem, StructItem};
use ruast::*;

/// Generate common methods for the struct
/// new, lock, deep_clone
pub(crate) fn generate_common_methods(struct_info: &StructItem) -> Vec<Item<AssocItemKind>> {
    let mut items = vec![];
    let impl_items: Vec<_> = struct_info.associated_impls().collect();
    let existing_new_fn =  impl_items.iter()
        .filter_map(|impl_item| impl_item.get_function("new"))
        .next();

    if let Some(existing_fn) = existing_new_fn {
        let params: Vec<Param> = existing_fn
            .sig()
            .inputs
            .iter()
            .map(|(name, _)| {
                Param::new(
                    Pat::ident(name),
                    Type::Path(Path::single("xlsx").chain(struct_info.name())),
                )
            })
            .collect();
        let args = params
            .iter()
            .map(|p| Expr::from(ExprKind::Path(Path::single(p.pat.to_string()))))
            .collect();

        let mut new_fn = Item::from(Fn::simple(
            "new",
            FnDecl::regular(
                params,
                Some(Type::Path(Path::single(PathSegment::new(
                    struct_info.name(),
                    None,
                )))),
            ),
            Block::single(ExprKind::Struct(Struct::new(
                Path::single(struct_info.name()),
                vec![ExprField::new(
                    "inner",
                    ExprKind::from(ExprKind::call(
                        ExprKind::Path(Path::single("Arc").chain("new")),
                        vec![Expr::from(ExprKind::call(
                            ExprKind::Path(Path::single("Mutex").chain("new")),
                            vec![Expr::from(ExprKind::call(
                                ExprKind::Path(
                                    Path::single("xlsx").chain(struct_info.name()).chain("new"),
                                ),
                                args,
                            ))],
                        ))],
                    )),
                )],
            ))),
        ));

        new_fn.vis = Visibility::Public;
        if let Some(doc) = &existing_fn.item().docs {
            new_fn.add_attr(Attribute::doc_comment(process_doc_comment(doc)));
        }
        // Add wasm_bindgen constructor attribute
        new_fn.add_attr(Attribute::normal(AttributeItem::new(
            "wasm_bindgen",
            AttrArgs::Delimited(DelimArgs::parenthesis(
                vec![Token::Ident("constructor".to_string())].into_tokens(),
            )),
        )));

        items.push(new_fn);
    }

    let mut lock_fn = Item::from(Fn::simple(
        "lock",
        FnDecl::regular(
            vec![Param::ref_self()],
            Some(Type::poly_path(
                "MutexGuard",
                vec![GenericArg::Type(Type::Path(
                    Path::single("xlsx").chain(struct_info.name()),
                ))],
            )),
        ),
        Block::single(ExprKind::method_call0(
            ExprKind::from(ExprKind::method_call0(
                ExprKind::from(ExprKind::field(
                    ExprKind::Path(Path::single("self")),
                    "inner",
                )),
                "lock",
            )),
            "unwrap",
        )),
    ));

    lock_fn.vis = Visibility::crate_();
    lock_fn.add_attr(new_line());
    items.push(lock_fn);

    let mut deep_clone_fn = Item::from(Fn::simple(
        "deep_clone",
        FnDecl::regular(
            vec![Param::ref_self()],
            Some(Type::Path(Path::single(PathSegment::new(
                struct_info.name(),
                None,
            )))),
        ),
        Block::new(
            vec![
                Stmt::Local(Local::new(
                    "inner",
                    None,
                    ExprKind::method_call0(
                        ExprKind::from(ExprKind::method_call0(
                            ExprKind::from(ExprKind::field(
                                ExprKind::Path(Path::single("self")),
                                "inner",
                            )),
                            "lock",
                        )),
                        "unwrap",
                    ),
                )),
                Stmt::from(Expr::from(Struct::new(
                    Path::single(struct_info.name()),
                    vec![ExprField::new(
                        "inner",
                        ExprKind::from(ExprKind::call(
                            ExprKind::Path(Path::single("Arc").chain("new")),
                            vec![Expr::from(ExprKind::call(
                                ExprKind::Path(Path::single("Mutex").chain("new")),
                                vec![Expr::from(ExprKind::method_call0(
                                    ExprKind::Path(Path::single("inner")),
                                    "clone",
                                ))],
                            ))],
                        )),
                    )],
                ))),
            ],
            None,
        ),
    ));

    deep_clone_fn.vis = Visibility::Public;
    deep_clone_fn.add_attr(new_line());
    deep_clone_fn.add_attr(Attribute::doc_comment(format!(
        "/// Deep clones a {} object.",
        struct_info.name()
    )));
    deep_clone_fn.add_attr(Attribute::normal(AttributeItem::new(
        "wasm_bindgen",
        AttrArgs::Delimited(DelimArgs::parenthesis(
            vec![
                Token::Ident("js_name".to_string()),
                Token::DocComment(" ".to_string()),
                Token::Eq,
                Token::DocComment(" ".to_string()),
                Token::Lit(Lit::str("clone")),
            ]
            .into_tokens(),
        )),
    )));
    items.push(deep_clone_fn);

    items
}

/// Generate wrapper methods for the original struct's methods and functions

pub(crate) fn generate_wrapper_methods(struct_info: &StructItem) -> Vec<Item<AssocItemKind>> {
    let impls = struct_info.associated_impls().collect::<Vec<_>>();

    let all_methods: Vec<_> = impls.iter()
        .flat_map(|impl_item| impl_item.functions().filter(|f| f.is_method()).collect::<Vec<_>>())
        .collect();

    let all_functions: Vec<_> = impls.iter()
        .flat_map(|impl_item| impl_item.functions().filter(|f| !f.is_method() && f.name() != "new").collect::<Vec<_>>())
        .collect();

    // Map methods and functions to wrapper methods
    let wrapped_methods = all_methods.into_iter()
        .map(|method| create_wrapper_method(&method, struct_info));
    let wrapped_functions = all_functions.into_iter()
        .map(|method| create_wrapper_method(&method, struct_info));

    wrapped_functions.chain(wrapped_methods).collect()
}

/// Create a wrapper method for a single function
fn create_wrapper_method(function: &FunctionItem, struct_info: &StructItem) -> Item<AssocItemKind> {
    let sig = function.sig();
    let generics = function.generics();
    let mut params = Vec::new();
    let mut call_args = Vec::new();

    // For methods, the first parameter is always a reference to self
    if function.is_method() {
        params.push(Param::ref_self());
    }

    for (i, (name, ty)) in sig.inputs.iter().enumerate() {
        // Skip the first parameter for methods (self)
        if function.is_method() && i == 0 {
            continue;
        }

        params.push(Param::new(Pat::ident(name), determine_param_type(ty, generics)));

        let arg_expr = match ty {
            rustdoc_types::Type::ResolvedPath(_) => Expr::from(ExprKind::method_call0(
                ExprKind::Path(Path::single(name)),
                "into",
            )),
            rustdoc_types::Type::BorrowedRef{ type_, .. } => {
                if let rustdoc_types::Type::ResolvedPath(_) = &**type_ {
                    Expr::from(ExprKind::method_call0(
                        ExprKind::Path(Path::single(name)),
                        "into",
                    ))
                } else {
                    Expr::from(ExprKind::Path(Path::single(name)))
                }
            }
            _ => {
                // For other types, just pass the name
                Expr::from(ExprKind::Path(Path::single(name)))
            }
        };
        call_args.push(arg_expr);
    }

    let method_body_macro = if function.is_method() {
        // For methods, create a self.method_name(args) expression
        let receiver = Expr::from(Path::single("self"));
        let method_segment = PathSegment::simple(function.name());
        let args: Vec<Expr> = call_args
            .iter()
            .map(|arg| Expr::from(Path::single(arg.to_string())))
            .collect();

        let method_call = MethodCall::new(receiver, method_segment, args);
        TokenStream::from(method_call)
    } else {
        // For functions, just use the arguments
        let mut tokens = Vec::new();
        for (i, arg) in call_args.iter().enumerate() {
            if i > 0 {
                tokens.push(Token::Comma);
            }
            tokens.push(Token::Ident(arg.to_string()));
        }
        tokens.into_tokens()
    };

    let method_body = if function.is_method() {
        Block::new(
            vec![Stmt::Semi(Semi(Expr::from(MacCall::new(
                Path::single("impl_method"),
                DelimArgs::parenthesis(method_body_macro),
            ))))],
            None,
        )
    } else {
        Block::new(
            vec![
                Stmt::Local(Local::simple(
                    "result",
                    ExprKind::Call(Call::new(
                        Path::single("xlsx")
                            .chain(struct_info.name())
                            .chain(function.name()),
                        vec![],
                    )),
                )),
                Stmt::Expr(Expr::from(ExprKind::Struct(Struct::new(
                    struct_info.name(),
                    vec![ExprField::new(
                        "inner",
                        ExprKind::from(ExprKind::call(
                            ExprKind::Path(Path::single("Arc").chain("new")),
                            vec![Expr::from(ExprKind::call(
                                ExprKind::Path(Path::single("Mutex").chain("new")),
                                vec![Expr::from(ExprKind::Path(Path::single("result")))],
                            ))],
                        )),
                    )],
                )))),
            ],
            None,
        )
    };

    // Create the function
    let mut wrapper_fn = Item::from(Fn::simple(
        function.name(),
        FnDecl::regular(
            params,
            Some(Type::Path(Path::single(PathSegment::new(
                struct_info.name(),
                None,
            )))),
        ),
        method_body,
    ));

    wrapper_fn.vis = Visibility::Public;
    wrapper_fn.add_attr(new_line());
    // inherit documentation if available
    if let Some(doc) = &function.item().docs {
        wrapper_fn.add_attr(Attribute::doc_comment(process_doc_comment(doc)));
    }
    // Add wasm_bindgen attribute with js_name in camelCase
    let js_name = to_camel_case(function.name());
    wrapper_fn.add_attr(Attribute::normal(AttributeItem::new(
        "wasm_bindgen",
        AttrArgs::Delimited(DelimArgs::parenthesis(
            vec![
                Token::Ident("js_name".to_string()),
                Token::DocComment(" ".to_string()),
                Token::Eq,
                Token::DocComment(" ".to_string()),
                Token::Lit(Lit::str(&js_name)),
                Token::Comma,
                Token::DocComment(" ".to_string()),
                Token::Ident("skip_jsdoc".to_string()),
            ]
            .into_tokens(),
        )),
    )));

    wrapper_fn
}

/// Determine the parameter type recursively
fn determine_param_type(ty: &rustdoc_types::Type, generics: &rustdoc_types::Generics) -> Type {
    let extract_into_inner = |path: &Vec<rustdoc_types::GenericBound>| {
        if let Some(rustdoc_types::GenericBound::TraitBound { trait_, .. }) = path.first() {
            if trait_.path == "Into" || trait_.path.ends_with("::Into") {
                if let Some(args) = &trait_.args {
                    if let rustdoc_types::GenericArgs::AngleBracketed { args, .. } = &**args {
                        if args.len() == 1 {
                            if let rustdoc_types::GenericArg::Type(inner_type) = &args[0] {
                                return determine_param_type(inner_type, generics);
                            }
                        }
                    }
                }
            }
        }
        panic!("Unsupported type for Into: {:?}", path);
    };
    match ty {
        rustdoc_types::Type::ResolvedPath(path) => Type::Path(Path::single(&path.path)),
        rustdoc_types::Type::Primitive(prim) => Type::Path(Path::single(prim)),
        rustdoc_types::Type::BorrowedRef { type_, .. } => determine_param_type(type_, generics),
        rustdoc_types::Type::Array { type_, .. } => {
            let inner_type = determine_param_type(type_, generics);
            Type::poly_path("Vec", vec![GenericArg::Type(inner_type)])
        }
        rustdoc_types::Type::ImplTrait(impl_trait) => {
            extract_into_inner(impl_trait)
        }
        rustdoc_types::Type::Generic(name) => {
            let generic_param = generics.params.iter()
                .find(|param| param.name == *name)
                .expect("Generic type not found in generics");
            
            for pred in  generics.where_predicates.iter() {
                if let rustdoc_types::WherePredicate::BoundPredicate{  type_, bounds , .. } = pred {
                    if let rustdoc_types::Type::Generic(generic_name) = type_ {
                        if  *generic_name == *name {
                            return extract_into_inner(bounds);
                        }
                    }
                }
            }

            let param_bounds = match &generic_param.kind {
                rustdoc_types::GenericParamDefKind::Type{bounds, ..} => bounds,
                _ => panic!("Unsupported generic parameter type: {:?}", generic_param),
            };
            extract_into_inner(param_bounds)
        }
        _ => panic!("Unsupported type: {:?}", ty),
    }
}
