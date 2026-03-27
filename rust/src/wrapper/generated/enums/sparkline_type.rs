use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `SparklineType` enum defines {@link Sparkline} types.
///
/// This is used with the {@link Sparkline#setType}(Sparkline::set_type())
/// method.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum SparklineType {
    /// A line style sparkline. This is the default.
    Line,
    /// A histogram style sparkline.
    Column,
    /// A positive/negative style sparkline. It looks similar to a histogram but
    /// all the bars are the same height,
    WinLose,
}

impl From<SparklineType> for xlsx::SparklineType {
    fn from(value: SparklineType) -> xlsx::SparklineType {
        match value {
            SparklineType::Line => xlsx::SparklineType::Line,
            SparklineType::Column => xlsx::SparklineType::Column,
            SparklineType::WinLose => xlsx::SparklineType::WinLose,
        }
    }
}
