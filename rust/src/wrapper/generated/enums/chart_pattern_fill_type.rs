use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ChartPatternFillType` enum defines the Chart pattern fill types.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ChartPatternFillType {
    /// Dotted 5 percent - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_5_percent.png">
    Dotted5Percent,
    /// Dotted 10 percent - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_10_percent.png">
    Dotted10Percent,
    /// Dotted 20 percent - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_20_percent.png">
    Dotted20Percent,
    /// Dotted 25 percent - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_25_percent.png">
    Dotted25Percent,
    /// Dotted 30 percent - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_30_percent.png">
    Dotted30Percent,
    /// Dotted 40 percent - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_40_percent.png">
    Dotted40Percent,
    /// Dotted 50 percent - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_50_percent.png">
    Dotted50Percent,
    /// Dotted 60 percent - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_60_percent.png">
    Dotted60Percent,
    /// Dotted 70 percent - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_70_percent.png">
    Dotted70Percent,
    /// Dotted 75 percent - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_75_percent.png">
    Dotted75Percent,
    /// Dotted 80 percent - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_80_percent.png">
    Dotted80Percent,
    /// Dotted 90 percent - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_90_percent.png">
    Dotted90Percent,
    /// Diagonal stripes light downwards - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_diagonal_stripes_light_downwards.png">
    DiagonalStripesLightDownwards,
    /// Diagonal stripes light upwards - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_diagonal_stripes_light_upwards.png">
    DiagonalStripesLightUpwards,
    /// Diagonal stripes dark downwards - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_diagonal_stripes_dark_downwards.png">
    DiagonalStripesDarkDownwards,
    /// Diagonal stripes dark upwards - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_diagonal_stripes_dark_upwards.png">
    DiagonalStripesDarkUpwards,
    /// Diagonal stripes wide downwards - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_diagonal_stripes_wide_downwards.png">
    DiagonalStripesWideDownwards,
    /// Diagonal stripes wide upwards - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_diagonal_stripes_wide_upwards.png">
    DiagonalStripesWideUpwards,
    /// Vertical stripes light - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_vertical_stripes_light.png">
    VerticalStripesLight,
    /// Horizontal stripes light - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_horizontal_stripes_light.png">
    HorizontalStripesLight,
    /// Vertical stripes narrow - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_vertical_stripes_narrow.png">
    VerticalStripesNarrow,
    /// Horizontal stripes narrow - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_horizontal_stripes_narrow.png">
    HorizontalStripesNarrow,
    /// Vertical stripes dark - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_vertical_stripes_dark.png">
    VerticalStripesDark,
    /// Horizontal stripes dark - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_horizontal_stripes_dark.png">
    HorizontalStripesDark,
    /// Stripes backslashes - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_stripes_backslashes.png">
    StripesBackslashes,
    /// Stripes forward slashes - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_stripes_forward_slashes.png">
    StripesForwardSlashes,
    /// Horizontal stripes alternating - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_horizontal_stripes_alternating.png">
    HorizontalStripesAlternating,
    /// Vertical stripes alternating - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_vertical_stripes_alternating.png">
    VerticalStripesAlternating,
    /// Small confetti - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_small_confetti.png">
    SmallConfetti,
    /// Large confetti - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_large_confetti.png">
    LargeConfetti,
    /// Zigzag - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_zigzag.png">
    Zigzag,
    /// Wave - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_wave.png">
    Wave,
    /// Diagonal brick - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_diagonal_brick.png">
    DiagonalBrick,
    /// Horizontal brick - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_horizontal_brick.png">
    HorizontalBrick,
    /// Weave - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_weave.png">
    Weave,
    /// Plaid - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_plaid.png">
    Plaid,
    /// Divot - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_divot.png">
    Divot,
    /// Dotted grid - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_grid.png">
    DottedGrid,
    /// Dotted diamond - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_diamond.png">
    DottedDiamond,
    /// Shingle - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_shingle.png">
    Shingle,
    /// Trellis - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_trellis.png">
    Trellis,
    /// Sphere - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_sphere.png">
    Sphere,
    /// Small grid - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_small_grid.png">
    SmallGrid,
    /// Large grid - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_large_grid.png">
    LargeGrid,
    /// Small checkerboard - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_small_checkerboard.png">
    SmallCheckerboard,
    /// Large checkerboard - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_large_checkerboard.png">
    LargeCheckerboard,
    /// Outlined diamond grid - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_outlined_diamond_grid.png">
    OutlinedDiamondGrid,
    /// Solid diamond grid - chart fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_solid_diamond_grid.png">
    SolidDiamondGrid,
}

impl From<ChartPatternFillType> for xlsx::ChartPatternFillType {
    fn from(value: ChartPatternFillType) -> xlsx::ChartPatternFillType {
        match value {
            ChartPatternFillType::Dotted5Percent => xlsx::ChartPatternFillType::Dotted5Percent,
            ChartPatternFillType::Dotted10Percent => xlsx::ChartPatternFillType::Dotted10Percent,
            ChartPatternFillType::Dotted20Percent => xlsx::ChartPatternFillType::Dotted20Percent,
            ChartPatternFillType::Dotted25Percent => xlsx::ChartPatternFillType::Dotted25Percent,
            ChartPatternFillType::Dotted30Percent => xlsx::ChartPatternFillType::Dotted30Percent,
            ChartPatternFillType::Dotted40Percent => xlsx::ChartPatternFillType::Dotted40Percent,
            ChartPatternFillType::Dotted50Percent => xlsx::ChartPatternFillType::Dotted50Percent,
            ChartPatternFillType::Dotted60Percent => xlsx::ChartPatternFillType::Dotted60Percent,
            ChartPatternFillType::Dotted70Percent => xlsx::ChartPatternFillType::Dotted70Percent,
            ChartPatternFillType::Dotted75Percent => xlsx::ChartPatternFillType::Dotted75Percent,
            ChartPatternFillType::Dotted80Percent => xlsx::ChartPatternFillType::Dotted80Percent,
            ChartPatternFillType::Dotted90Percent => xlsx::ChartPatternFillType::Dotted90Percent,
            ChartPatternFillType::DiagonalStripesLightDownwards => {
                xlsx::ChartPatternFillType::DiagonalStripesLightDownwards
            }
            ChartPatternFillType::DiagonalStripesLightUpwards => {
                xlsx::ChartPatternFillType::DiagonalStripesLightUpwards
            }
            ChartPatternFillType::DiagonalStripesDarkDownwards => {
                xlsx::ChartPatternFillType::DiagonalStripesDarkDownwards
            }
            ChartPatternFillType::DiagonalStripesDarkUpwards => {
                xlsx::ChartPatternFillType::DiagonalStripesDarkUpwards
            }
            ChartPatternFillType::DiagonalStripesWideDownwards => {
                xlsx::ChartPatternFillType::DiagonalStripesWideDownwards
            }
            ChartPatternFillType::DiagonalStripesWideUpwards => {
                xlsx::ChartPatternFillType::DiagonalStripesWideUpwards
            }
            ChartPatternFillType::VerticalStripesLight => {
                xlsx::ChartPatternFillType::VerticalStripesLight
            }
            ChartPatternFillType::HorizontalStripesLight => {
                xlsx::ChartPatternFillType::HorizontalStripesLight
            }
            ChartPatternFillType::VerticalStripesNarrow => {
                xlsx::ChartPatternFillType::VerticalStripesNarrow
            }
            ChartPatternFillType::HorizontalStripesNarrow => {
                xlsx::ChartPatternFillType::HorizontalStripesNarrow
            }
            ChartPatternFillType::VerticalStripesDark => {
                xlsx::ChartPatternFillType::VerticalStripesDark
            }
            ChartPatternFillType::HorizontalStripesDark => {
                xlsx::ChartPatternFillType::HorizontalStripesDark
            }
            ChartPatternFillType::StripesBackslashes => {
                xlsx::ChartPatternFillType::StripesBackslashes
            }
            ChartPatternFillType::StripesForwardSlashes => {
                xlsx::ChartPatternFillType::StripesForwardSlashes
            }
            ChartPatternFillType::HorizontalStripesAlternating => {
                xlsx::ChartPatternFillType::HorizontalStripesAlternating
            }
            ChartPatternFillType::VerticalStripesAlternating => {
                xlsx::ChartPatternFillType::VerticalStripesAlternating
            }
            ChartPatternFillType::SmallConfetti => xlsx::ChartPatternFillType::SmallConfetti,
            ChartPatternFillType::LargeConfetti => xlsx::ChartPatternFillType::LargeConfetti,
            ChartPatternFillType::Zigzag => xlsx::ChartPatternFillType::Zigzag,
            ChartPatternFillType::Wave => xlsx::ChartPatternFillType::Wave,
            ChartPatternFillType::DiagonalBrick => xlsx::ChartPatternFillType::DiagonalBrick,
            ChartPatternFillType::HorizontalBrick => xlsx::ChartPatternFillType::HorizontalBrick,
            ChartPatternFillType::Weave => xlsx::ChartPatternFillType::Weave,
            ChartPatternFillType::Plaid => xlsx::ChartPatternFillType::Plaid,
            ChartPatternFillType::Divot => xlsx::ChartPatternFillType::Divot,
            ChartPatternFillType::DottedGrid => xlsx::ChartPatternFillType::DottedGrid,
            ChartPatternFillType::DottedDiamond => xlsx::ChartPatternFillType::DottedDiamond,
            ChartPatternFillType::Shingle => xlsx::ChartPatternFillType::Shingle,
            ChartPatternFillType::Trellis => xlsx::ChartPatternFillType::Trellis,
            ChartPatternFillType::Sphere => xlsx::ChartPatternFillType::Sphere,
            ChartPatternFillType::SmallGrid => xlsx::ChartPatternFillType::SmallGrid,
            ChartPatternFillType::LargeGrid => xlsx::ChartPatternFillType::LargeGrid,
            ChartPatternFillType::SmallCheckerboard => {
                xlsx::ChartPatternFillType::SmallCheckerboard
            }
            ChartPatternFillType::LargeCheckerboard => {
                xlsx::ChartPatternFillType::LargeCheckerboard
            }
            ChartPatternFillType::OutlinedDiamondGrid => {
                xlsx::ChartPatternFillType::OutlinedDiamondGrid
            }
            ChartPatternFillType::SolidDiamondGrid => xlsx::ChartPatternFillType::SolidDiamondGrid,
        }
    }
}
