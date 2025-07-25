#[macro_export]
macro_rules! impl_method {
    ($self:ident.$method:ident($($arg:expr),*)) => {
        let mut lock = $self.lock();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.$method($($arg),*);
        let _ = std::mem::replace(&mut *lock, inner);
        return Self {
            inner: Arc::clone(&$self.inner),
        }
    };
}
