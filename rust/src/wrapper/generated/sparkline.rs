use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `Sparkline` struct is used to create an object to represent a sparkline
/// that can be inserted into a worksheet.
///
/// Sparklines are a feature of Excel 2010+ which allows you to add small charts
/// to worksheet cells. These are useful for showing data trends in a compact
/// visual format.
///
/// The following example was used to generate the above file.
///
/// In Excel sparklines can be added as a single entity in a cell that refers to
/// a 1D data range or as a "group" sparkline that is applied across a 1D range
/// and refers to data in a 2D range. A grouped sparkline uses one sparkline for
/// the specified range and any changes to it are applied to the entire
/// sparkline group.
///
/// The {@link Worksheet#addSparkline} method
/// shown allows you to add a sparkline to a single cell that displays data from
/// a 1D range of cells whereas the
/// {@link Worksheet#addSparklineGroup})
/// method applies the group sparkline to a range.
#[derive(Clone)]
#[wasm_bindgen]
pub struct Sparkline {
    pub(crate) inner: Arc<Mutex<xlsx::Sparkline>>,
}

#[wasm_bindgen]
impl Sparkline {
    #[wasm_bindgen(constructor)]
    pub fn new() -> Sparkline {
        Sparkline {
            inner: Arc::new(Mutex::new(xlsx::Sparkline::new())),
        }
    }
}
