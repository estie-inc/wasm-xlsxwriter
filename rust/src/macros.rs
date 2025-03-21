#[macro_export]
macro_rules! wrap_struct {
    ($struct_name:ident,
     $inner_type:ty,
     new($($new_arg_name:ident: $new_arg_type:ty),*),
     $($method:ident($($arg_name:ident: $arg_type:ty),*)),*) => {
        #[wasm_bindgen]
        #[derive(Clone)]
        pub struct $struct_name {
            pub(crate) inner: std::sync::Arc<std::sync::Mutex<$inner_type>>,
        }

        #[wasm_bindgen]
        impl $struct_name {
            /// Lock the inner value and return a reference.
            pub(crate) fn lock(&self) -> std::sync::MutexGuard<'_, $inner_type> {
                self.inner.lock().unwrap()
            }

            /// Create a new instance.
            #[wasm_bindgen(constructor)]
            pub fn new($($new_arg_name: $new_arg_type),*) -> Self {
                Self {
                    inner: std::sync::Arc::new(std::sync::Mutex::new(<$inner_type>::new($($new_arg_name),*))),
                }
            }

            #[wasm_bindgen(js_name = clone)]
            pub fn deep_clone(&self) -> Self {
                let inner = self.lock();
                Self {
                    inner: std::sync::Arc::new(std::sync::Mutex::new(inner.clone())),
                }
            }

            $(
                #[wasm_bindgen(js_name = $method)]
                pub fn $method(&self, $($arg_name: $arg_type),*) -> Self {
                    let mut lock = self.lock();
                    let mut inner = lock.clone();
                    inner = inner.$method($($arg_name),*);
                    *lock = inner;
                    Self {
                        inner: std::sync::Arc::clone(&self.inner),
                    }
                }
            )*
        }
    };
} 