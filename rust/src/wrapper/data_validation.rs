use std::sync::Arc;

use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;
use wasm_bindgen::prelude::*;

use crate::wrapper::{DataValidation, WasmResult};

/// Validation rule with numeric values.
/// Used for allow_whole_number (i32), allow_decimal_number (f64),
/// allow_text_length (u32), and allow_time (f64).
#[derive(Debug, Clone, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum DataValidationRuleNumber {
    Between(f64, f64),
    NotBetween(f64, f64),
    EqualTo(f64),
    NotEqualTo(f64),
    GreaterThan(f64),
    GreaterThanOrEqualTo(f64),
    LessThan(f64),
    LessThanOrEqualTo(f64),
}

/// Validation rule with string values.
/// Used for formula variants (string → Formula) and date variants (string → ExcelDateTime).
#[derive(Debug, Clone, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum DataValidationRuleString {
    Between(String, String),
    NotBetween(String, String),
    EqualTo(String),
    NotEqualTo(String),
    GreaterThan(String),
    GreaterThanOrEqualTo(String),
    LessThan(String),
    LessThanOrEqualTo(String),
}

macro_rules! map_number_rule {
    ($rule:expr, $ty:ty) => {
        match $rule {
            DataValidationRuleNumber::Between(a, b) => {
                xlsx::DataValidationRule::Between(a as $ty, b as $ty)
            }
            DataValidationRuleNumber::NotBetween(a, b) => {
                xlsx::DataValidationRule::NotBetween(a as $ty, b as $ty)
            }
            DataValidationRuleNumber::EqualTo(v) => xlsx::DataValidationRule::EqualTo(v as $ty),
            DataValidationRuleNumber::NotEqualTo(v) => {
                xlsx::DataValidationRule::NotEqualTo(v as $ty)
            }
            DataValidationRuleNumber::GreaterThan(v) => {
                xlsx::DataValidationRule::GreaterThan(v as $ty)
            }
            DataValidationRuleNumber::GreaterThanOrEqualTo(v) => {
                xlsx::DataValidationRule::GreaterThanOrEqualTo(v as $ty)
            }
            DataValidationRuleNumber::LessThan(v) => xlsx::DataValidationRule::LessThan(v as $ty),
            DataValidationRuleNumber::LessThanOrEqualTo(v) => {
                xlsx::DataValidationRule::LessThanOrEqualTo(v as $ty)
            }
        }
    };
}

fn to_formula_rule(
    rule: DataValidationRuleString,
) -> xlsx::DataValidationRule<xlsx::Formula> {
    match rule {
        DataValidationRuleString::Between(a, b) => {
            xlsx::DataValidationRule::Between(xlsx::Formula::new(a), xlsx::Formula::new(b))
        }
        DataValidationRuleString::NotBetween(a, b) => {
            xlsx::DataValidationRule::NotBetween(xlsx::Formula::new(a), xlsx::Formula::new(b))
        }
        DataValidationRuleString::EqualTo(v) => {
            xlsx::DataValidationRule::EqualTo(xlsx::Formula::new(v))
        }
        DataValidationRuleString::NotEqualTo(v) => {
            xlsx::DataValidationRule::NotEqualTo(xlsx::Formula::new(v))
        }
        DataValidationRuleString::GreaterThan(v) => {
            xlsx::DataValidationRule::GreaterThan(xlsx::Formula::new(v))
        }
        DataValidationRuleString::GreaterThanOrEqualTo(v) => {
            xlsx::DataValidationRule::GreaterThanOrEqualTo(xlsx::Formula::new(v))
        }
        DataValidationRuleString::LessThan(v) => {
            xlsx::DataValidationRule::LessThan(xlsx::Formula::new(v))
        }
        DataValidationRuleString::LessThanOrEqualTo(v) => {
            xlsx::DataValidationRule::LessThanOrEqualTo(xlsx::Formula::new(v))
        }
    }
}

fn to_datetime_rule(
    rule: DataValidationRuleString,
) -> Result<xlsx::DataValidationRule<xlsx::ExcelDateTime>, xlsx::XlsxError> {
    Ok(match rule {
        DataValidationRuleString::Between(a, b) => xlsx::DataValidationRule::Between(
            xlsx::ExcelDateTime::parse_from_str(&a)?,
            xlsx::ExcelDateTime::parse_from_str(&b)?,
        ),
        DataValidationRuleString::NotBetween(a, b) => xlsx::DataValidationRule::NotBetween(
            xlsx::ExcelDateTime::parse_from_str(&a)?,
            xlsx::ExcelDateTime::parse_from_str(&b)?,
        ),
        DataValidationRuleString::EqualTo(v) => {
            xlsx::DataValidationRule::EqualTo(xlsx::ExcelDateTime::parse_from_str(&v)?)
        }
        DataValidationRuleString::NotEqualTo(v) => {
            xlsx::DataValidationRule::NotEqualTo(xlsx::ExcelDateTime::parse_from_str(&v)?)
        }
        DataValidationRuleString::GreaterThan(v) => {
            xlsx::DataValidationRule::GreaterThan(xlsx::ExcelDateTime::parse_from_str(&v)?)
        }
        DataValidationRuleString::GreaterThanOrEqualTo(v) => {
            xlsx::DataValidationRule::GreaterThanOrEqualTo(
                xlsx::ExcelDateTime::parse_from_str(&v)?,
            )
        }
        DataValidationRuleString::LessThan(v) => {
            xlsx::DataValidationRule::LessThan(xlsx::ExcelDateTime::parse_from_str(&v)?)
        }
        DataValidationRuleString::LessThanOrEqualTo(v) => {
            xlsx::DataValidationRule::LessThanOrEqualTo(
                xlsx::ExcelDateTime::parse_from_str(&v)?,
            )
        }
    })
}

fn return_self(this: &DataValidation) -> DataValidation {
    DataValidation {
        inner: Arc::clone(&this.inner),
    }
}

#[wasm_bindgen]
impl DataValidation {
    #[wasm_bindgen(js_name = "allowListStrings")]
    pub fn allow_list_strings(&self, list: Vec<String>) -> WasmResult<DataValidation> {
        let str_refs: Vec<&str> = list.iter().map(|s| s.as_str()).collect();
        let mut lock = self.inner.lock().unwrap();
        let inner = std::mem::replace(&mut *lock, xlsx::DataValidation::new());
        *lock = inner.allow_list_strings(&str_refs)?;
        Ok(return_self(self))
    }

    #[wasm_bindgen(js_name = "allowWholeNumber")]
    pub fn allow_whole_number(&self, rule: DataValidationRuleNumber) -> DataValidation {
        let r = map_number_rule!(rule, i32);
        let mut lock = self.inner.lock().unwrap();
        let inner = std::mem::replace(&mut *lock, xlsx::DataValidation::new());
        *lock = inner.allow_whole_number(r);
        return_self(self)
    }

    #[wasm_bindgen(js_name = "allowWholeNumberFormula")]
    pub fn allow_whole_number_formula(
        &self,
        rule: DataValidationRuleString,
    ) -> DataValidation {
        let r = to_formula_rule(rule);
        let mut lock = self.inner.lock().unwrap();
        let inner = std::mem::replace(&mut *lock, xlsx::DataValidation::new());
        *lock = inner.allow_whole_number_formula(r);
        return_self(self)
    }

    #[wasm_bindgen(js_name = "allowDecimalNumber")]
    pub fn allow_decimal_number(&self, rule: DataValidationRuleNumber) -> DataValidation {
        let r = map_number_rule!(rule, f64);
        let mut lock = self.inner.lock().unwrap();
        let inner = std::mem::replace(&mut *lock, xlsx::DataValidation::new());
        *lock = inner.allow_decimal_number(r);
        return_self(self)
    }

    #[wasm_bindgen(js_name = "allowDecimalNumberFormula")]
    pub fn allow_decimal_number_formula(
        &self,
        rule: DataValidationRuleString,
    ) -> DataValidation {
        let r = to_formula_rule(rule);
        let mut lock = self.inner.lock().unwrap();
        let inner = std::mem::replace(&mut *lock, xlsx::DataValidation::new());
        *lock = inner.allow_decimal_number_formula(r);
        return_self(self)
    }

    #[wasm_bindgen(js_name = "allowTextLength")]
    pub fn allow_text_length(&self, rule: DataValidationRuleNumber) -> DataValidation {
        let r = map_number_rule!(rule, u32);
        let mut lock = self.inner.lock().unwrap();
        let inner = std::mem::replace(&mut *lock, xlsx::DataValidation::new());
        *lock = inner.allow_text_length(r);
        return_self(self)
    }

    #[wasm_bindgen(js_name = "allowTextLengthFormula")]
    pub fn allow_text_length_formula(
        &self,
        rule: DataValidationRuleString,
    ) -> DataValidation {
        let r = to_formula_rule(rule);
        let mut lock = self.inner.lock().unwrap();
        let inner = std::mem::replace(&mut *lock, xlsx::DataValidation::new());
        *lock = inner.allow_text_length_formula(r);
        return_self(self)
    }

    #[wasm_bindgen(js_name = "allowDate")]
    pub fn allow_date(&self, rule: DataValidationRuleString) -> WasmResult<DataValidation> {
        let r = to_datetime_rule(rule)?;
        let mut lock = self.inner.lock().unwrap();
        let inner = std::mem::replace(&mut *lock, xlsx::DataValidation::new());
        *lock = inner.allow_date(r);
        Ok(return_self(self))
    }

    #[wasm_bindgen(js_name = "allowDateFormula")]
    pub fn allow_date_formula(
        &self,
        rule: DataValidationRuleString,
    ) -> DataValidation {
        let r = to_formula_rule(rule);
        let mut lock = self.inner.lock().unwrap();
        let inner = std::mem::replace(&mut *lock, xlsx::DataValidation::new());
        *lock = inner.allow_date_formula(r);
        return_self(self)
    }

    #[wasm_bindgen(js_name = "allowTime")]
    pub fn allow_time(&self, rule: DataValidationRuleString) -> WasmResult<DataValidation> {
        let r = to_datetime_rule(rule)?;
        let mut lock = self.inner.lock().unwrap();
        let inner = std::mem::replace(&mut *lock, xlsx::DataValidation::new());
        *lock = inner.allow_time(r);
        Ok(return_self(self))
    }

    #[wasm_bindgen(js_name = "allowTimeFormula")]
    pub fn allow_time_formula(
        &self,
        rule: DataValidationRuleString,
    ) -> DataValidation {
        let r = to_formula_rule(rule);
        let mut lock = self.inner.lock().unwrap();
        let inner = std::mem::replace(&mut *lock, xlsx::DataValidation::new());
        *lock = inner.allow_time_formula(r);
        return_self(self)
    }
}
