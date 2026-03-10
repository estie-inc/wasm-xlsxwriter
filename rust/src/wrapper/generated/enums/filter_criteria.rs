use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `FilterCriteria` enum defines logical filter criteria used in an
/// autofilter.
///
/// These filter criteria are used with the FilterCondition
/// add_custom_filter()(FilterCondition::add_custom_filter) method.
///
/// Currently only Excel's string and number filter operations are supported.
/// The numeric style criteria such as `>=` can also be applied to strings (like
/// in Rust) but the string operations like `BeginsWith` are only applied to
/// strings in Excel.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum FilterCriteria {
    /// Show numbers or strings that are equal to the filter value.
    EqualTo,
    /// Show numbers or strings that are not equal to the filter value.
    NotEqualTo,
    /// Show numbers or strings that are greater than the filter value.
    GreaterThan,
    /// Show numbers or strings that are greater than or equal to the filter value.
    GreaterThanOrEqualTo,
    /// Show numbers or strings that are less than the filter value.
    LessThan,
    /// Show numbers or strings that are less than or equal to the filter value.
    LessThanOrEqualTo,
    /// Show strings that begin with the filter string value.
    BeginsWith,
    /// Show strings that do not begin with the filter string value.
    DoesNotBeginWith,
    /// Show strings that end with the filter string value.
    EndsWith,
    /// Show strings that do not end with the filter string value.
    DoesNotEndWith,
    /// Show strings that contain with the filter string value.
    Contains,
    /// Show strings that do not contain with the filter string value.
    DoesNotContain,
}

impl From<FilterCriteria> for xlsx::FilterCriteria {
    fn from(value: FilterCriteria) -> xlsx::FilterCriteria {
        match value {
            FilterCriteria::EqualTo => xlsx::FilterCriteria::EqualTo,
            FilterCriteria::NotEqualTo => xlsx::FilterCriteria::NotEqualTo,
            FilterCriteria::GreaterThan => xlsx::FilterCriteria::GreaterThan,
            FilterCriteria::GreaterThanOrEqualTo => xlsx::FilterCriteria::GreaterThanOrEqualTo,
            FilterCriteria::LessThan => xlsx::FilterCriteria::LessThan,
            FilterCriteria::LessThanOrEqualTo => xlsx::FilterCriteria::LessThanOrEqualTo,
            FilterCriteria::BeginsWith => xlsx::FilterCriteria::BeginsWith,
            FilterCriteria::DoesNotBeginWith => xlsx::FilterCriteria::DoesNotBeginWith,
            FilterCriteria::EndsWith => xlsx::FilterCriteria::EndsWith,
            FilterCriteria::DoesNotEndWith => xlsx::FilterCriteria::DoesNotEndWith,
            FilterCriteria::Contains => xlsx::FilterCriteria::Contains,
            FilterCriteria::DoesNotContain => xlsx::FilterCriteria::DoesNotContain,
        }
    }
}
