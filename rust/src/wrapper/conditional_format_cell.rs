use std::sync::Arc;

use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;
use wasm_bindgen::prelude::*;

use crate::wrapper::ConditionalFormatCell;

/// Rule for cell conditional formatting.
/// Values are strings: numbers as "42", text as "hello", formulas as "=A1".
/// Internally converted to ConditionalFormatValue which handles type detection.
#[derive(Debug, Clone, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ConditionalFormatCellRule {
    Between(String, String),
    NotBetween(String, String),
    EqualTo(String),
    NotEqualTo(String),
    GreaterThan(String),
    GreaterThanOrEqualTo(String),
    LessThan(String),
    LessThanOrEqualTo(String),
}

fn to_xlsx_rule(
    rule: ConditionalFormatCellRule,
) -> xlsx::ConditionalFormatCellRule<xlsx::ConditionalFormatValue> {
    match rule {
        ConditionalFormatCellRule::Between(a, b) => xlsx::ConditionalFormatCellRule::Between(
            xlsx::ConditionalFormatValue::from(a),
            xlsx::ConditionalFormatValue::from(b),
        ),
        ConditionalFormatCellRule::NotBetween(a, b) => {
            xlsx::ConditionalFormatCellRule::NotBetween(
                xlsx::ConditionalFormatValue::from(a),
                xlsx::ConditionalFormatValue::from(b),
            )
        }
        ConditionalFormatCellRule::EqualTo(v) => {
            xlsx::ConditionalFormatCellRule::EqualTo(
                xlsx::ConditionalFormatValue::from(v),
            )
        }
        ConditionalFormatCellRule::NotEqualTo(v) => {
            xlsx::ConditionalFormatCellRule::NotEqualTo(
                xlsx::ConditionalFormatValue::from(v),
            )
        }
        ConditionalFormatCellRule::GreaterThan(v) => {
            xlsx::ConditionalFormatCellRule::GreaterThan(
                xlsx::ConditionalFormatValue::from(v),
            )
        }
        ConditionalFormatCellRule::GreaterThanOrEqualTo(v) => {
            xlsx::ConditionalFormatCellRule::GreaterThanOrEqualTo(
                xlsx::ConditionalFormatValue::from(v),
            )
        }
        ConditionalFormatCellRule::LessThan(v) => {
            xlsx::ConditionalFormatCellRule::LessThan(
                xlsx::ConditionalFormatValue::from(v),
            )
        }
        ConditionalFormatCellRule::LessThanOrEqualTo(v) => {
            xlsx::ConditionalFormatCellRule::LessThanOrEqualTo(
                xlsx::ConditionalFormatValue::from(v),
            )
        }
    }
}

#[wasm_bindgen]
impl ConditionalFormatCell {
    #[wasm_bindgen(js_name = "setRule")]
    pub fn set_rule(&self, rule: ConditionalFormatCellRule) -> ConditionalFormatCell {
        let mut lock = self.inner.lock().unwrap();
        let inner =
            std::mem::replace(&mut *lock, xlsx::ConditionalFormatCell::new());
        *lock = inner.set_rule(to_xlsx_rule(rule));
        ConditionalFormatCell {
            inner: Arc::clone(&self.inner),
        }
    }
}
