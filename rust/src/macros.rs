macro_rules! wrap_struct {
    ($struct_name:ident, $inner_type:ty, $( $method:ident ( $($arg_name:ident : $arg_type:ty),* ) ),* ) => {
        use std::sync::{Arc, Mutex};

        pub struct $struct_name {
            inner: Arc<Mutex<$inner_type>>,
        }

        #[wasm_bindgen]
        #[derive(Clone)]
        impl $struct_name {
            /// Lock the inner value and return a reference.
            pub(crate) fn lock(&self) -> std::sync::MutexGuard<'_, $inner_type> {
                self.inner.lock().unwrap()
            }

            /// Create a new instance of `$struct_name`.
            ///
            /// @param {string} link - A string representing a URL.
            /// @returns {Self} - The new `$struct_name` instance.
            #[wasm_bindgen(constructor, skip_jsdoc)]
            pub fn new($($new_arg_name: $new_arg_type),*) -> Self {
                Self {
                    inner: Arc::new(Mutex::new(<$inner_type>::new(link))),
                }
            }

            /// Clone the object deeply.
            #[wasm_bindgen(js_name = "clone")]
            pub fn deep_clone(&self) -> Self {
                let inner = self.lock();
                Self {
                    inner: Arc::new(Mutex::new(inner.clone())),
                }
            }

            $(
                #[wasm_bindgen(js_name = concat!(
                    // convert to snake case
                    let mut name = stringify!($method)
                        .split('_')
                        .map(|word| {
                            let mut chars = word.chars();
                            match chars.next() {
                                Some(first) => first.to_ascii_uppercase().to_string() + chars.as_str(),
                                None => "".to_string(),
                            }
                        })
                        .collect::<String>();
                    name[..1].make_ascii_lowercase();
                    name
                ))]
                pub fn $method(&self, $($arg_name: $arg_type),*) -> Self {
                    let mut lock = self.lock();
                    let mut inner = std::mem::replace(&mut *lock, <$inner_type>::new(""));
                    inner = inner.$method($($arg_name),*);
                    let _ = std::mem::replace(&mut *lock, inner);
                    Self {
                        inner: Arc::clone(&self.inner),
                    }
                }
            )*
        }
    };
}
