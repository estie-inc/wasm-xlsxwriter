use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `FormatBorder` enum defines the Excel border types that can be added to
/// a {@link Format} pattern.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify, Default)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum FormatBorder {
    /// No border.
    #[default]
    None,
    /// Thin border style.
    Thin,
    /// Medium border style.
    Medium,
    /// Dashed border style.
    Dashed,
    /// Dotted border style.
    Dotted,
    /// Thick border style.
    Thick,
    /// Double border style.
    Double,
    /// Hair border style.
    Hair,
    /// Medium dashed border style.
    MediumDashed,
    /// Dash-dot border style.
    DashDot,
    /// Medium dash-dot border style.
    MediumDashDot,
    /// Dash-dot-dot border style.
    DashDotDot,
    /// Medium dash-dot-dot border style.
    MediumDashDotDot,
    /// Slant dash-dot border style.
    SlantDashDot,
}

impl From<FormatBorder> for xlsx::FormatBorder {
    fn from(value: FormatBorder) -> xlsx::FormatBorder {
        match value {
            FormatBorder::None => xlsx::FormatBorder::None,
            FormatBorder::Thin => xlsx::FormatBorder::Thin,
            FormatBorder::Medium => xlsx::FormatBorder::Medium,
            FormatBorder::Dashed => xlsx::FormatBorder::Dashed,
            FormatBorder::Dotted => xlsx::FormatBorder::Dotted,
            FormatBorder::Thick => xlsx::FormatBorder::Thick,
            FormatBorder::Double => xlsx::FormatBorder::Double,
            FormatBorder::Hair => xlsx::FormatBorder::Hair,
            FormatBorder::MediumDashed => xlsx::FormatBorder::MediumDashed,
            FormatBorder::DashDot => xlsx::FormatBorder::DashDot,
            FormatBorder::MediumDashDot => xlsx::FormatBorder::MediumDashDot,
            FormatBorder::DashDotDot => xlsx::FormatBorder::DashDotDot,
            FormatBorder::MediumDashDotDot => xlsx::FormatBorder::MediumDashDotDot,
            FormatBorder::SlantDashDot => xlsx::FormatBorder::SlantDashDot,
        }
    }
}
