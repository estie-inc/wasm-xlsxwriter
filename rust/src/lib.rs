#![allow(clippy::new_without_default)]

mod error;
mod wrapper;

trait DeepClone {
    fn deep_clone(&self) -> Self;
}
