use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `SparklineType` enum defines Sparkline types.
///
/// This is used with the Sparkline.setType()(Sparkline::set_type())
/// method.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum SparklineType {
    /// A line style sparkline. This is the default.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/sparkline_type_line.png">
    Line,
    /// A histogram style sparkline.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/sparkline_type_column.png">
    Column,
    /// A positive/negative style sparkline. It looks similar to a histogram but
    /// all the bars are the same height,
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/sparkline_type_winlose.png">
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
