use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ChartArea` struct is a representation of the background area object of
/// an Excel chart.
///
/// The `ChartArea` struct can be used to configure properties of the chart area
/// such as the formatting and is usually obtained via the
/// {@link Chart#chartArea} method.
///
/// It is used in conjunction with the {@link Chart} struct.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartArea {
    pub(crate) parent: Arc<Mutex<xlsx::Chart>>,
}

#[wasm_bindgen]
impl ChartArea {}
