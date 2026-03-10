use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ShapePatternFillType` enum defines the Shape pattern fill types.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ShapePatternFillType {
    /// Dotted 5 percent - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_5_percent.png">
    Dotted5Percent,
    /// Dotted 10 percent - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_10_percent.png">
    Dotted10Percent,
    /// Dotted 20 percent - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_20_percent.png">
    Dotted20Percent,
    /// Dotted 25 percent - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_25_percent.png">
    Dotted25Percent,
    /// Dotted 30 percent - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_30_percent.png">
    Dotted30Percent,
    /// Dotted 40 percent - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_40_percent.png">
    Dotted40Percent,
    /// Dotted 50 percent - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_50_percent.png">
    Dotted50Percent,
    /// Dotted 60 percent - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_60_percent.png">
    Dotted60Percent,
    /// Dotted 70 percent - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_70_percent.png">
    Dotted70Percent,
    /// Dotted 75 percent - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_75_percent.png">
    Dotted75Percent,
    /// Dotted 80 percent - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_80_percent.png">
    Dotted80Percent,
    /// Dotted 90 percent - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_90_percent.png">
    Dotted90Percent,
    /// Diagonal stripes light downwards - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_diagonal_stripes_light_downwards.png">
    DiagonalStripesLightDownwards,
    /// Diagonal stripes light upwards - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_diagonal_stripes_light_upwards.png">
    DiagonalStripesLightUpwards,
    /// Diagonal stripes dark downwards - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_diagonal_stripes_dark_downwards.png">
    DiagonalStripesDarkDownwards,
    /// Diagonal stripes dark upwards - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_diagonal_stripes_dark_upwards.png">
    DiagonalStripesDarkUpwards,
    /// Diagonal stripes wide downwards - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_diagonal_stripes_wide_downwards.png">
    DiagonalStripesWideDownwards,
    /// Diagonal stripes wide upwards - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_diagonal_stripes_wide_upwards.png">
    DiagonalStripesWideUpwards,
    /// Vertical stripes light - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_vertical_stripes_light.png">
    VerticalStripesLight,
    /// Horizontal stripes light - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_horizontal_stripes_light.png">
    HorizontalStripesLight,
    /// Vertical stripes narrow - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_vertical_stripes_narrow.png">
    VerticalStripesNarrow,
    /// Horizontal stripes narrow - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_horizontal_stripes_narrow.png">
    HorizontalStripesNarrow,
    /// Vertical stripes dark - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_vertical_stripes_dark.png">
    VerticalStripesDark,
    /// Horizontal stripes dark - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_horizontal_stripes_dark.png">
    HorizontalStripesDark,
    /// Stripes backslashes - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_stripes_backslashes.png">
    StripesBackslashes,
    /// Stripes forward slashes - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_stripes_forward_slashes.png">
    StripesForwardSlashes,
    /// Horizontal stripes alternating - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_horizontal_stripes_alternating.png">
    HorizontalStripesAlternating,
    /// Vertical stripes alternating - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_vertical_stripes_alternating.png">
    VerticalStripesAlternating,
    /// Small confetti - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_small_confetti.png">
    SmallConfetti,
    /// Large confetti - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_large_confetti.png">
    LargeConfetti,
    /// Zigzag - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_zigzag.png">
    Zigzag,
    /// Wave - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_wave.png">
    Wave,
    /// Diagonal brick - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_diagonal_brick.png">
    DiagonalBrick,
    /// Horizontal brick - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_horizontal_brick.png">
    HorizontalBrick,
    /// Weave - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_weave.png">
    Weave,
    /// Plaid - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_plaid.png">
    Plaid,
    /// Divot - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_divot.png">
    Divot,
    /// Dotted grid - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_grid.png">
    DottedGrid,
    /// Dotted diamond - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_diamond.png">
    DottedDiamond,
    /// Shingle - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_shingle.png">
    Shingle,
    /// Trellis - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_trellis.png">
    Trellis,
    /// Sphere - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_sphere.png">
    Sphere,
    /// Small grid - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_small_grid.png">
    SmallGrid,
    /// Large grid - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_large_grid.png">
    LargeGrid,
    /// Small checkerboard - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_small_checkerboard.png">
    SmallCheckerboard,
    /// Large checkerboard - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_large_checkerboard.png">
    LargeCheckerboard,
    /// Outlined diamond grid - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_outlined_diamond_grid.png">
    OutlinedDiamondGrid,
    /// Solid diamond grid - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_solid_diamond_grid.png">
    SolidDiamondGrid,
}

impl From<ShapePatternFillType> for xlsx::ShapePatternFillType {
    fn from(value: ShapePatternFillType) -> xlsx::ShapePatternFillType {
        match value {
            ShapePatternFillType::Dotted5Percent => xlsx::ShapePatternFillType::Dotted5Percent,
            ShapePatternFillType::Dotted10Percent => xlsx::ShapePatternFillType::Dotted10Percent,
            ShapePatternFillType::Dotted20Percent => xlsx::ShapePatternFillType::Dotted20Percent,
            ShapePatternFillType::Dotted25Percent => xlsx::ShapePatternFillType::Dotted25Percent,
            ShapePatternFillType::Dotted30Percent => xlsx::ShapePatternFillType::Dotted30Percent,
            ShapePatternFillType::Dotted40Percent => xlsx::ShapePatternFillType::Dotted40Percent,
            ShapePatternFillType::Dotted50Percent => xlsx::ShapePatternFillType::Dotted50Percent,
            ShapePatternFillType::Dotted60Percent => xlsx::ShapePatternFillType::Dotted60Percent,
            ShapePatternFillType::Dotted70Percent => xlsx::ShapePatternFillType::Dotted70Percent,
            ShapePatternFillType::Dotted75Percent => xlsx::ShapePatternFillType::Dotted75Percent,
            ShapePatternFillType::Dotted80Percent => xlsx::ShapePatternFillType::Dotted80Percent,
            ShapePatternFillType::Dotted90Percent => xlsx::ShapePatternFillType::Dotted90Percent,
            ShapePatternFillType::DiagonalStripesLightDownwards => {
                xlsx::ShapePatternFillType::DiagonalStripesLightDownwards
            }
            ShapePatternFillType::DiagonalStripesLightUpwards => {
                xlsx::ShapePatternFillType::DiagonalStripesLightUpwards
            }
            ShapePatternFillType::DiagonalStripesDarkDownwards => {
                xlsx::ShapePatternFillType::DiagonalStripesDarkDownwards
            }
            ShapePatternFillType::DiagonalStripesDarkUpwards => {
                xlsx::ShapePatternFillType::DiagonalStripesDarkUpwards
            }
            ShapePatternFillType::DiagonalStripesWideDownwards => {
                xlsx::ShapePatternFillType::DiagonalStripesWideDownwards
            }
            ShapePatternFillType::DiagonalStripesWideUpwards => {
                xlsx::ShapePatternFillType::DiagonalStripesWideUpwards
            }
            ShapePatternFillType::VerticalStripesLight => {
                xlsx::ShapePatternFillType::VerticalStripesLight
            }
            ShapePatternFillType::HorizontalStripesLight => {
                xlsx::ShapePatternFillType::HorizontalStripesLight
            }
            ShapePatternFillType::VerticalStripesNarrow => {
                xlsx::ShapePatternFillType::VerticalStripesNarrow
            }
            ShapePatternFillType::HorizontalStripesNarrow => {
                xlsx::ShapePatternFillType::HorizontalStripesNarrow
            }
            ShapePatternFillType::VerticalStripesDark => {
                xlsx::ShapePatternFillType::VerticalStripesDark
            }
            ShapePatternFillType::HorizontalStripesDark => {
                xlsx::ShapePatternFillType::HorizontalStripesDark
            }
            ShapePatternFillType::StripesBackslashes => {
                xlsx::ShapePatternFillType::StripesBackslashes
            }
            ShapePatternFillType::StripesForwardSlashes => {
                xlsx::ShapePatternFillType::StripesForwardSlashes
            }
            ShapePatternFillType::HorizontalStripesAlternating => {
                xlsx::ShapePatternFillType::HorizontalStripesAlternating
            }
            ShapePatternFillType::VerticalStripesAlternating => {
                xlsx::ShapePatternFillType::VerticalStripesAlternating
            }
            ShapePatternFillType::SmallConfetti => xlsx::ShapePatternFillType::SmallConfetti,
            ShapePatternFillType::LargeConfetti => xlsx::ShapePatternFillType::LargeConfetti,
            ShapePatternFillType::Zigzag => xlsx::ShapePatternFillType::Zigzag,
            ShapePatternFillType::Wave => xlsx::ShapePatternFillType::Wave,
            ShapePatternFillType::DiagonalBrick => xlsx::ShapePatternFillType::DiagonalBrick,
            ShapePatternFillType::HorizontalBrick => xlsx::ShapePatternFillType::HorizontalBrick,
            ShapePatternFillType::Weave => xlsx::ShapePatternFillType::Weave,
            ShapePatternFillType::Plaid => xlsx::ShapePatternFillType::Plaid,
            ShapePatternFillType::Divot => xlsx::ShapePatternFillType::Divot,
            ShapePatternFillType::DottedGrid => xlsx::ShapePatternFillType::DottedGrid,
            ShapePatternFillType::DottedDiamond => xlsx::ShapePatternFillType::DottedDiamond,
            ShapePatternFillType::Shingle => xlsx::ShapePatternFillType::Shingle,
            ShapePatternFillType::Trellis => xlsx::ShapePatternFillType::Trellis,
            ShapePatternFillType::Sphere => xlsx::ShapePatternFillType::Sphere,
            ShapePatternFillType::SmallGrid => xlsx::ShapePatternFillType::SmallGrid,
            ShapePatternFillType::LargeGrid => xlsx::ShapePatternFillType::LargeGrid,
            ShapePatternFillType::SmallCheckerboard => {
                xlsx::ShapePatternFillType::SmallCheckerboard
            }
            ShapePatternFillType::LargeCheckerboard => {
                xlsx::ShapePatternFillType::LargeCheckerboard
            }
            ShapePatternFillType::OutlinedDiamondGrid => {
                xlsx::ShapePatternFillType::OutlinedDiamondGrid
            }
            ShapePatternFillType::SolidDiamondGrid => xlsx::ShapePatternFillType::SolidDiamondGrid,
        }
    }
}
