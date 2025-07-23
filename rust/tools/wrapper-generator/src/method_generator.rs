use crate::utils::{add_doc_comment_marker, omit_after_example, to_camel_case};
use crate_info_extractor::{ImplFnInfo, StructInfo};
use ruast::*;

/// Generate wrapper methods for the original struct's methods and functions
pub fn generate_wrapper_methods(struct_info: &StructInfo) -> Vec<Item<AssocItemKind>> {
    let mut items = Vec::new();

    for function in &struct_info.methods {
        items.push(create_wrapper_method(function, struct_info));
    }

    for function in &struct_info.functions {
        // skip new because it is already handled in impl_common_methods
        if function.name == "new" {
            continue;
        }
        items.push(create_wrapper_method(function, struct_info));
    }

    // // Add the impl_function macro definition if there are any functions
    // if !struct_info.functions.is_empty() {
    //     let macro_def = create_impl_function_macro(struct_info);
    //     items.push(macro_def);
    // }

    items
}

/// Create the impl_function macro definition
fn create_impl_function_macro(struct_info: &StructInfo) -> Item<ItemKind> {
    // Create a macro definition for impl_function
    let macro_body = r#"
        ($function:ident($($arg:expr),*)) => {
            let result = xlsx::$function($($arg),*);
            return $struct_name {
                inner: Arc::new(Mutex::new(result)),
            }
        };
    "#
    .replace("$struct_name", &struct_info.name);

    // Create a macro_rules item
    let macro_item = Item::from(MacroDef::new(
        "impl_function",
        DelimArgs::brace(
            vec![
                Token::Ident("function".to_string()),
                Token::Eq,
                Token::Lit(Lit::str(macro_body)),
            ]
            .into_tokens(),
        ),
    ));

    macro_item
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

    // Create the method body
    // Use different macro based on whether it's a method or function
    let macro_name = if function.is_method {
        "impl_method"
    } else {
        "impl_function"
    };

    // For methods, create a self.method_name(args) expression
    let macro_content = if function.is_method {
        // Create a method call expression using AST: self.method_name(args)
        let receiver = Expr::from(Path::single("self"));
        let method_segment = PathSegment::simple(function.name.clone());

        // Convert call_args to Expr
        let args: Vec<Expr> = call_args
            .iter()
            .map(|arg| Expr::from(Path::single(arg.to_string())))
            .collect();

        // Create the method call
        let method_call = MethodCall::new(receiver, method_segment, args);

        // Convert to token stream
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
        Path::single(macro_name),
        DelimArgs::parenthesis(macro_content),
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

    // Add static keyword for functions (not methods)
    if !function.is_method {
        wrapper_fn.add_attr(Attribute::normal(AttributeItem::new(
            "wasm_bindgen",
            AttrArgs::Delimited(DelimArgs::parenthesis(
                vec![
                    Token::Ident("static_method_of".to_string()),
                    Token::Eq,
                    Token::Lit(Lit::str(&struct_info.name)),
                ]
                .into_tokens(),
            )),
        )));
    }

    // Set visibility to public
    wrapper_fn.vis = Visibility::Public;

    // Add documentation if available
    if let Some(doc) = &function.doc {
        wrapper_fn.add_attr(Attribute::doc_comment(add_doc_comment_marker(
            omit_after_example(doc),
        )));
    }

    // Add wasm_bindgen attribute with js_name
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

/// Check if a type needs .into()
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
