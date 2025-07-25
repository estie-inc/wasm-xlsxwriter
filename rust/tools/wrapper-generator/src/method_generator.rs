use crate::utils::{new_line, process_doc_comment, to_camel_case};
use crate_info_extractor::{ImplFnInfo, StructInfo};
use ruast::*;
use std::process::abort;

/// Generate common methods for the struct
/// new, lock, deep_clone
pub(crate) fn generate_common_methods(struct_info: &StructInfo) -> Vec<Item<AssocItemKind>> {
    let mut items = vec![];

    let existing_new_fn = struct_info.functions.iter().find(|f| f.name == "new");
    if let Some(existing_fn) = existing_new_fn {
        let params = if let Some(sig) = &existing_fn.sig {
            sig.inputs
                .iter()
                .map(|(name, _)| {
                    Param::new(
                        Pat::ident(name),
                        Type::Path(Path::single("xlsx").chain(struct_info.name.clone())),
                    )
                })
                .collect()
        } else {
            vec![]
        };
        let args = params
            .iter()
            .map(|p| Expr::from(ExprKind::Path(Path::single(p.pat.to_string()))))
            .collect();

        let mut new_fn = Item::from(Fn::simple(
            "new",
            FnDecl::regular(
                params,
                Some(Type::Path(Path::single(PathSegment::new(
                    &struct_info.name,
                    None,
                )))),
            ),
            Block::single(ExprKind::Struct(Struct::new(
                Path::single(&struct_info.name),
                vec![ExprField::new(
                    "inner",
                    ExprKind::from(ExprKind::call(
                        ExprKind::Path(Path::single("Arc").chain("new")),
                        vec![Expr::from(ExprKind::call(
                            ExprKind::Path(Path::single("Mutex").chain("new")),
                            vec![Expr::from(ExprKind::call(
                                ExprKind::Path(
                                    Path::single("xlsx")
                                        .chain(struct_info.name.clone())
                                        .chain("new"),
                                ),
                                args,
                            ))],
                        ))],
                    )),
                )],
            ))),
        ));

        new_fn.vis = Visibility::Public;
        if let Some(doc) = &existing_fn.doc {
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
                    Path::single("xlsx").chain(struct_info.name.clone()),
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
                &struct_info.name,
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
                    Path::single(&struct_info.name),
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
        struct_info.name
    )));
    deep_clone_fn.add_attr(Attribute::normal(AttributeItem::new(
        "wasm_bindgen",
        AttrArgs::Delimited(DelimArgs::parenthesis(
            vec![
                Token::Ident("js_name".to_string()),
                Token::Eq,
                Token::Lit(Lit::str("clone")),
            ]
            .into_tokens(),
        )),
    )));
    items.push(deep_clone_fn);

    items
}

/// Generate wrapper methods for the original struct's methods and functions

pub(crate) fn generate_wrapper_methods(struct_info: &StructInfo) -> Vec<Item<AssocItemKind>> {
    let wrapped_methods = struct_info
        .methods
        .iter()
        .map(|method| create_wrapper_method(method, struct_info));
    let wrapped_functions = struct_info
        .functions
        .iter()
        .filter(|f| f.name != "new") // Skip the new function
        .map(|function| create_wrapper_method(function, struct_info));

    wrapped_functions.chain(wrapped_methods).collect()
}

/// Create a wrapper method for a single function
fn create_wrapper_method(function: &ImplFnInfo, struct_info: &StructInfo) -> Item<AssocItemKind> {
    let sig = function.sig.as_ref().unwrap();
    let mut params = Vec::new();
    let mut call_args = Vec::new();

    // For methods, the first parameter is always a reference to self
    if function.is_method {
        params.push(Param::ref_self());
    }

    for (i, (name, ty)) in sig.inputs.iter().enumerate() {
        // Skip the first parameter for methods (self)
        if function.is_method && i == 0 {
            continue;
        }

        params.push(Param::new(Pat::ident(name), determine_param_type(ty)));

        let arg_expr = match ty {
            rustdoc_types::Type::ResolvedPath(_) => Expr::from(ExprKind::method_call0(
                ExprKind::Path(Path::single(name)),
                "into",
            )),
            _ => {
                // For other types, just pass the name
                Expr::from(ExprKind::Path(Path::single(name)))
            }
        };
        call_args.push(arg_expr);
    }

    let method_body_macro = if function.is_method {
        // For methods, create a self.method_name(args) expression
        let receiver = Expr::from(Path::single("self"));
        let method_segment = PathSegment::simple(function.name.clone());
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
    let method_body = Block::single(ExprKind::MacCall(MacCall::new(
        Path::single(if function.is_method {
            "impl_method"
        } else {
            "impl_function"
        }),
        DelimArgs::parenthesis(method_body_macro),
    )));

    // Create the function
    let mut wrapper_fn = Item::from(Fn::simple(
        &function.name,
        FnDecl::regular(
            params,
            Some(Type::Path(Path::single(PathSegment::new(
                &struct_info.name,
                None,
            )))),
        ),
        method_body,
    ));

    wrapper_fn.vis = Visibility::Public;
    wrapper_fn.add_attr(new_line());
    // inherit documentation if available
    if let Some(doc) = &function.doc {
        wrapper_fn.add_attr(Attribute::doc_comment(process_doc_comment(doc)));
    }
    // Add wasm_bindgen attribute with js_name in camelCase
    let js_name = to_camel_case(&function.name);
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
fn determine_param_type(ty: &rustdoc_types::Type) -> Type {
    match ty {
        rustdoc_types::Type::ResolvedPath(path) => Type::Path(Path::single(&path.path)),
        rustdoc_types::Type::Primitive(prim) => Type::Path(Path::single(prim)),
        rustdoc_types::Type::BorrowedRef { type_, .. } => determine_param_type(type_),
        rustdoc_types::Type::Array { type_, .. } => {
            let inner_type = determine_param_type(type_);
            Type::poly_path("Vec", vec![GenericArg::Type(inner_type)])
        }
        rustdoc_types::Type::ImplTrait(impl_trait) => {
            if impl_trait.len() == 1 {
                if let rustdoc_types::GenericBound::TraitBound { trait_, .. } = &impl_trait[0] {
                    if trait_.path == "Into" || trait_.path.ends_with("::Into") {
                        if let Some(args) = &trait_.args {
                            if let rustdoc_types::GenericArgs::AngleBracketed { args, .. } = &**args
                            {
                                if args.len() == 1 {
                                    if let rustdoc_types::GenericArg::Type(inner_type) = &args[0] {
                                        return determine_param_type(inner_type);
                                    }
                                }
                            }
                        }
                    }
                }
            }

            abort()
        }
        rustdoc_types::Type::Generic(_) => {
            // TODO: Extract the generic type from constraints if available
            Type::Path(Path::single("UnknownGeneric"))
        }
        _ => abort(),
    }
}
