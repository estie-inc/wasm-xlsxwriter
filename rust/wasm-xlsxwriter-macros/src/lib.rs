use proc_macro::TokenStream;
use proc_macro2::TokenStream as TokenStream2;
use quote::quote;
use syn::{parse_macro_input, ItemStruct, Ident, Type, parse::Parse, Token};

struct WrapperArgs {
    inner_type: Type,
    dummy_fn: Option<Ident>,
}

impl Parse for WrapperArgs {
    fn parse(input: syn::parse::ParseStream) -> syn::Result<Self> {
        let inner_type: Type = input.parse()?;

        let dummy_fn = if input.peek(Token![,]) {
            input.parse::<Token![,]>()?;
            Some(input.parse()?)
        } else {
            None
        };

        Ok(WrapperArgs {
            inner_type,
            dummy_fn,
        })
    }
}

/// A procedural macro to generate wrapper code for wasm-xlsxwriter.
/// 
/// This macro generates the appropriate wrapper code for both Default and non-Default types.
/// It generates the struct definition and the impl_method macro, but does not generate any methods
/// to avoid conflicts with existing methods.
/// 
/// # Usage
/// 
/// ```
/// #[wasm_wrapper(xlsx::Format)]
/// pub struct Format;
/// ```
/// 
/// For types that don't implement Default, you need to provide a function that creates a dummy value:
/// 
/// ```
/// #[wasm_wrapper(xlsx::Image, new_dummy_image)]
/// pub struct Image;
/// ```
/// 
/// The macro generates:
/// 1. A struct with an Arc<Mutex<T>> field
/// 2. An impl_method macro for wrapping methods
/// 
/// You'll need to implement all methods manually using the impl_method macro.
// Function to generate wrapper methods for common patterns
fn generate_common_wrapper_methods(_struct_name: &Ident, _inner_type: &Type) -> TokenStream2 {
    // We won't generate any methods to avoid conflicts with existing methods
    // Users will need to implement all methods manually

    quote! {
        // No methods generated to avoid conflicts
    }
}

#[proc_macro_attribute]
pub fn wasm_wrapper(attr: TokenStream, item: TokenStream) -> TokenStream {
    let args = parse_macro_input!(attr as WrapperArgs);
    let input = parse_macro_input!(item as ItemStruct);

    let struct_name = &input.ident;
    let inner_type = &args.inner_type;
    let dummy_fn = &args.dummy_fn;

    // Generate the appropriate impl_method macro based on whether a dummy function is provided
    let impl_method_macro = if let Some(dummy_fn) = dummy_fn {
        // For types that don't implement Default
        quote! {
            macro_rules! impl_method {
                ($self:ident.$method:ident($($arg:expr),*)) => {
                    let mut lock = $self.inner.lock().unwrap();
                    let mut inner = std::mem::replace(
                        &mut *lock,
                        #dummy_fn(),
                    );
                    inner = inner.$method($($arg),*);
                    let _ = std::mem::replace(&mut *lock, inner);
                    return #struct_name {
                        inner: Arc::clone(&$self.inner),
                    }
                };
            }
        }
    } else {
        // For types that implement Default
        quote! {
            macro_rules! impl_method {
                ($self:ident.$method:ident($($arg:expr),*)) => {
                    let mut lock = $self.inner.lock().unwrap();
                    let mut inner = std::mem::take(&mut *lock);
                    inner = inner.$method($($arg),*);
                    let _ = std::mem::replace(&mut *lock, inner);
                    return #struct_name {
                        inner: Arc::clone(&$self.inner),
                    }
                };
            }
        }
    };

    // Generate common wrapper methods
    let wrapper_methods = generate_common_wrapper_methods(struct_name, inner_type);

    // Generate the output
    let output = quote! {
        #[derive(Clone)]
        #[wasm_bindgen]
        pub struct #struct_name {
            pub(crate) inner: Arc<Mutex<#inner_type>>,
        }

        #impl_method_macro

        #wrapper_methods
    };

    output.into()
}
