use crate::utils::{new_line, process_doc_comment, to_camel_case};
use crate_info_extractor::{ImplFnInfo, StructInfo};
use ruast::*;

/// Generate common methods for the struct
/// new, lock, deep_clone
pub(crate) fn generate_common_methods(struct_info: &StructInfo) -> Vec<Item<AssocItemKind>> {
    let mut items = vec![];

    let existing_new_fn = struct_info.functions.iter().find(|f| f.name == "new");
    if let Some(existing_fn) = existing_new_fn {
        let params = if let Some(sig) = &existing_fn.sig {
            sig.inputs
                .iter()
                .map(|(name, ty)| {
                    Param::new(
                        Pat::ident(name),
                        Type::Path(Path::single("xlsx").chain(struct_info.name.clone())),
                    )
                })
                .collect()
        } else {
            vec![]
        };

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
                            vec![Expr::from(ExprKind::Path(Path::single("inner")))],
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
    // Skip functions without signatures
    let sig = function.sig.as_ref().unwrap();

    // Prepare parameters
    let mut params = Vec::new();

    // Add &self parameter for methods, skip for functions
    if function.is_method {
        params.push(Param::ref_self());
    }

    // Process parameters
    let mut call_args = Vec::new();

    // Determine which parameters to process
    let start_idx = if function.is_method { 1 } else { 0 };

    for (i, (name, ty)) in sig.inputs.iter().enumerate().skip(start_idx) {
        // Extract the type name from the Type enum
        let type_name = match ty {
            rustdoc_types::Type::ResolvedPath(path) => path
                .path
                .split("::")
                .last()
                .unwrap_or(&path.path)
                .to_string(),
            rustdoc_types::Type::Primitive(prim) => prim.clone(),
            rustdoc_types::Type::BorrowedRef { type_, .. } => match &**type_ {
                rustdoc_types::Type::ResolvedPath(path) => path
                    .path
                    .split("::")
                    .last()
                    .unwrap_or(&path.path)
                    .to_string(),
                rustdoc_types::Type::Primitive(prim) => prim.clone(),
                _ => "Unknown".to_string(),
            },
            _ => "Unknown".to_string(),
        };

        // Create parameter with the appropriate type
        let param_type = determine_param_type(&type_name);

        params.push(Param::new(Pat::ident(name), param_type));

        // Create argument for the method call
        let arg_expr = if is_primitive_type(&type_name) {
            // Primitive types are passed directly
            Expr::from(ExprKind::Path(Path::single(name)))
        } else if needs_into(&type_name) {
            // Enum types that need .into()
            Expr::from(ExprKind::method_call0(
                ExprKind::Path(Path::single(name)),
                "into",
            ))
        } else {
            // Other types are passed directly
            Expr::from(ExprKind::Path(Path::single(name)))
        };

        call_args.push(arg_expr);
    }

    // Create the method body using the impl_method! macro
    let mut args_str = String::new();

    for (i, arg) in call_args.iter().enumerate() {
        if i > 0 {
            args_str.push_str(", ");
        }

        // This is a simplification - in a real implementation we'd need to convert the Expr to a string
        args_str.push_str(&format!("${}", i + 1));
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
                Token::Eq,
                Token::Lit(Lit::str(&js_name)),
                Token::Comma,
                Token::Ident("skip_jsdoc".to_string()),
            ]
            .into_tokens(),
        )),
    )));

    wrapper_fn
}

/// Determine the parameter type for the wrapper method

fn determine_param_type(original_type: &str) -> Type {
    if is_primitive_type(original_type) {
        // Primitive types are used as-is
        Type::Path(Path::single(original_type))
    } else if original_type.starts_with("Format") && original_type != "Format" {
        // Format* enum types (except Format itself) are used as-is
        Type::Path(Path::single(original_type))
    } else if original_type == "Color" {
        // Color is a wrapper type
        Type::Path(Path::single("Color"))
    } else if original_type.starts_with("xlsx::") {
        // Convert xlsx::Type to Type
        let wrapper_type = original_type
            .strip_prefix("xlsx::")
            .unwrap_or(original_type);

        Type::Path(Path::single(wrapper_type))
    } else {
        // Other types are used as-is
        Type::Path(Path::single(original_type))
    }
}

/// Check if a type is a primitive type

fn is_primitive_type(ty: &str) -> bool {
    matches!(
        ty,
        "i8" | "i16"
            | "i32"
            | "i64"
            | "i128"
            | "isize"
            | "u8"
            | "u16"
            | "u32"
            | "u64"
            | "u128"
            | "usize"
            | "f32"
            | "f64"
            | "bool"
            | "char"
            | "&str"
            | "String"
    )
}

/// Check if a type needs .into()\
fn needs_into(ty: &str) -> bool {
    // Enum types like FormatAlign, FormatBorder, etc. need .into()
    // We can identify them by their naming convention (Format* but not Format itself)
    ty.starts_with("Format") && ty != "Format"
}

/// Check if a type needs to access the inner field
fn needs_inner(ty: &str) -> bool {
    // Struct types like Color need to access the inner field
    matches!(ty, "Color")
}
