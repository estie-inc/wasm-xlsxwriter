use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ChartPatternFillType` enum defines the {@link Chart} pattern fill types.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ChartPatternFillType {
    /// Dotted 5 percent - chart fill pattern.
    Dotted5Percent,
    /// Dotted 10 percent - chart fill pattern.
    Dotted10Percent,
    /// Dotted 20 percent - chart fill pattern.
    Dotted20Percent,
    /// Dotted 25 percent - chart fill pattern.
    Dotted25Percent,
    /// Dotted 30 percent - chart fill pattern.
    Dotted30Percent,
    /// Dotted 40 percent - chart fill pattern.
    Dotted40Percent,
    /// Dotted 50 percent - chart fill pattern.
    Dotted50Percent,
    /// Dotted 60 percent - chart fill pattern.
    Dotted60Percent,
    /// Dotted 70 percent - chart fill pattern.
    Dotted70Percent,
    /// Dotted 75 percent - chart fill pattern.
    Dotted75Percent,
    /// Dotted 80 percent - chart fill pattern.
    Dotted80Percent,
    /// Dotted 90 percent - chart fill pattern.
    Dotted90Percent,
    /// Diagonal stripes light downwards - chart fill pattern.
    DiagonalStripesLightDownwards,
    /// Diagonal stripes light upwards - chart fill pattern.
    DiagonalStripesLightUpwards,
    /// Diagonal stripes dark downwards - chart fill pattern.
    DiagonalStripesDarkDownwards,
    /// Diagonal stripes dark upwards - chart fill pattern.
    DiagonalStripesDarkUpwards,
    /// Diagonal stripes wide downwards - chart fill pattern.
    DiagonalStripesWideDownwards,
    /// Diagonal stripes wide upwards - chart fill pattern.
    DiagonalStripesWideUpwards,
    /// Vertical stripes light - chart fill pattern.
    VerticalStripesLight,
    /// Horizontal stripes light - chart fill pattern.
    HorizontalStripesLight,
    /// Vertical stripes narrow - chart fill pattern.
    VerticalStripesNarrow,
    /// Horizontal stripes narrow - chart fill pattern.
    HorizontalStripesNarrow,
    /// Vertical stripes dark - chart fill pattern.
    VerticalStripesDark,
    /// Horizontal stripes dark - chart fill pattern.
    HorizontalStripesDark,
    /// Stripes backslashes - chart fill pattern.
    StripesBackslashes,
    /// Stripes forward slashes - chart fill pattern.
    StripesForwardSlashes,
    /// Horizontal stripes alternating - chart fill pattern.
    HorizontalStripesAlternating,
    /// Vertical stripes alternating - chart fill pattern.
    VerticalStripesAlternating,
    /// Small confetti - chart fill pattern.
    SmallConfetti,
    /// Large confetti - chart fill pattern.
    LargeConfetti,
    /// Zigzag - chart fill pattern.
    Zigzag,
    /// Wave - chart fill pattern.
    Wave,
    /// Diagonal brick - chart fill pattern.
    DiagonalBrick,
    /// Horizontal brick - chart fill pattern.
    HorizontalBrick,
    /// Weave - chart fill pattern.
    Weave,
    /// Plaid - chart fill pattern.
    Plaid,
    /// Divot - chart fill pattern.
    Divot,
    /// Dotted grid - chart fill pattern.
    DottedGrid,
    /// Dotted diamond - chart fill pattern.
    DottedDiamond,
    /// Shingle - chart fill pattern.
    Shingle,
    /// Trellis - chart fill pattern.
    Trellis,
    /// Sphere - chart fill pattern.
    Sphere,
    /// Small grid - chart fill pattern.
    SmallGrid,
    /// Large grid - chart fill pattern.
    LargeGrid,
    /// Small checkerboard - chart fill pattern.
    SmallCheckerboard,
    /// Large checkerboard - chart fill pattern.
    LargeCheckerboard,
    /// Outlined diamond grid - chart fill pattern.
    OutlinedDiamondGrid,
    /// Solid diamond grid - chart fill pattern.
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
