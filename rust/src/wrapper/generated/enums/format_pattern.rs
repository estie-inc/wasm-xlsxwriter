use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `FormatPattern` enum defines the Excel pattern types that can be added to
/// a Format.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify, Default)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum FormatPattern {
    /// Automatic or Empty pattern.
    #[default]
    None,
    /// Solid pattern.
    Solid,
    /// Medium gray pattern.
    MediumGray,
    /// Dark gray pattern.
    DarkGray,
    /// Light gray pattern.
    LightGray,
    /// Dark horizontal line pattern.
    DarkHorizontal,
    /// Dark vertical line pattern.
    DarkVertical,
    /// Dark diagonal stripe pattern.
    DarkDown,
    /// Reverse dark diagonal stripe pattern.
    DarkUp,
    /// Dark grid pattern.
    DarkGrid,
    /// Dark trellis pattern.
    DarkTrellis,
    /// Light horizontal Line pattern.
    LightHorizontal,
    /// Light vertical line pattern.
    LightVertical,
    /// Light diagonal stripe pattern.
    LightDown,
    /// Reverse light diagonal stripe pattern.
    LightUp,
    /// Light grid pattern.
    LightGrid,
    /// Light trellis pattern.
    LightTrellis,
    /// 12.5% gray pattern.
    Gray125,
    /// 6.25% gray pattern.
    Gray0625,
}

impl From<FormatPattern> for xlsx::FormatPattern {
    fn from(value: FormatPattern) -> xlsx::FormatPattern {
        match value {
            FormatPattern::None => xlsx::FormatPattern::None,
            FormatPattern::Solid => xlsx::FormatPattern::Solid,
            FormatPattern::MediumGray => xlsx::FormatPattern::MediumGray,
            FormatPattern::DarkGray => xlsx::FormatPattern::DarkGray,
            FormatPattern::LightGray => xlsx::FormatPattern::LightGray,
            FormatPattern::DarkHorizontal => xlsx::FormatPattern::DarkHorizontal,
            FormatPattern::DarkVertical => xlsx::FormatPattern::DarkVertical,
            FormatPattern::DarkDown => xlsx::FormatPattern::DarkDown,
            FormatPattern::DarkUp => xlsx::FormatPattern::DarkUp,
            FormatPattern::DarkGrid => xlsx::FormatPattern::DarkGrid,
            FormatPattern::DarkTrellis => xlsx::FormatPattern::DarkTrellis,
            FormatPattern::LightHorizontal => xlsx::FormatPattern::LightHorizontal,
            FormatPattern::LightVertical => xlsx::FormatPattern::LightVertical,
            FormatPattern::LightDown => xlsx::FormatPattern::LightDown,
            FormatPattern::LightUp => xlsx::FormatPattern::LightUp,
            FormatPattern::LightGrid => xlsx::FormatPattern::LightGrid,
            FormatPattern::LightTrellis => xlsx::FormatPattern::LightTrellis,
            FormatPattern::Gray125 => xlsx::FormatPattern::Gray125,
            FormatPattern::Gray0625 => xlsx::FormatPattern::Gray0625,
        }
    }
}
