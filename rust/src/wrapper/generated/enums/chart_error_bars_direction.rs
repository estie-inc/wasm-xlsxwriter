use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ChartErrorBarsDirection` enum defines the error bar direction for a
/// chart series {@link ChartErrorBars}.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ChartErrorBarsDirection {
    /// The error bars extend in both directions. This is the default.
    Both,
    /// The error bars extend in the negative direction only.
    Minus,
    /// The error bars extend in the positive direction only.
    Plus,
}

impl From<ChartErrorBarsDirection> for xlsx::ChartErrorBarsDirection {
    fn from(value: ChartErrorBarsDirection) -> xlsx::ChartErrorBarsDirection {
        match value {
            ChartErrorBarsDirection::Both => xlsx::ChartErrorBarsDirection::Both,
            ChartErrorBarsDirection::Minus => xlsx::ChartErrorBarsDirection::Minus,
            ChartErrorBarsDirection::Plus => xlsx::ChartErrorBarsDirection::Plus,
        }
    }
}
