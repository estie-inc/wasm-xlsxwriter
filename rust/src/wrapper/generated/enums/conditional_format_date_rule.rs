use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ConditionalFormatDateRule` enum defines the conditional format
/// criteria for {@link ConditionalFormatDate}.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ConditionalFormatDateRule {
    /// Show the conditional format for dates occurring yesterday. This is the
    /// default.
    Yesterday,
    /// Show the conditional format for dates occurring today, relative to when
    /// the file is opened.
    Today,
    /// Show the conditional format for dates occurring tomorrow, relative to
    /// when the file is opened.
    Tomorrow,
    /// Show the conditional format for dates occurring in the last 7 days,
    /// relative to when the file is opened.
    Last7Days,
    /// Show the conditional format for dates occurring in the last week,
    /// relative to when the file is opened.
    LastWeek,
    /// Show the conditional format for dates occurring this week, relative to
    /// when the file is opened.
    ThisWeek,
    /// Show the conditional format for dates occurring in the next week,
    /// relative to when the file is opened.
    NextWeek,
    /// Show the conditional format for dates occurring in the last month,
    /// relative to when the file is opened.
    LastMonth,
    /// Show the conditional format for dates occurring this month, relative to
    /// when the file is opened.
    ThisMonth,
    /// Show the conditional format for dates occurring in the next month,
    /// relative to when the file is opened.
    NextMonth,
}

impl From<ConditionalFormatDateRule> for xlsx::ConditionalFormatDateRule {
    fn from(value: ConditionalFormatDateRule) -> xlsx::ConditionalFormatDateRule {
        match value {
            ConditionalFormatDateRule::Yesterday => xlsx::ConditionalFormatDateRule::Yesterday,
            ConditionalFormatDateRule::Today => xlsx::ConditionalFormatDateRule::Today,
            ConditionalFormatDateRule::Tomorrow => xlsx::ConditionalFormatDateRule::Tomorrow,
            ConditionalFormatDateRule::Last7Days => xlsx::ConditionalFormatDateRule::Last7Days,
            ConditionalFormatDateRule::LastWeek => xlsx::ConditionalFormatDateRule::LastWeek,
            ConditionalFormatDateRule::ThisWeek => xlsx::ConditionalFormatDateRule::ThisWeek,
            ConditionalFormatDateRule::NextWeek => xlsx::ConditionalFormatDateRule::NextWeek,
            ConditionalFormatDateRule::LastMonth => xlsx::ConditionalFormatDateRule::LastMonth,
            ConditionalFormatDateRule::ThisMonth => xlsx::ConditionalFormatDateRule::ThisMonth,
            ConditionalFormatDateRule::NextMonth => xlsx::ConditionalFormatDateRule::NextMonth,
        }
    }
}
