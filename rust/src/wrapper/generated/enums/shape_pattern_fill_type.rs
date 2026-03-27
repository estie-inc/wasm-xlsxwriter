use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ShapePatternFillType` enum defines the {@link Shape} pattern fill types.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ShapePatternFillType {
    /// Dotted 5 percent - shape fill pattern.
    Dotted5Percent,
    /// Dotted 10 percent - shape fill pattern.
    Dotted10Percent,
    /// Dotted 20 percent - shape fill pattern.
    Dotted20Percent,
    /// Dotted 25 percent - shape fill pattern.
    Dotted25Percent,
    /// Dotted 30 percent - shape fill pattern.
    Dotted30Percent,
    /// Dotted 40 percent - shape fill pattern.
    Dotted40Percent,
    /// Dotted 50 percent - shape fill pattern.
    Dotted50Percent,
    /// Dotted 60 percent - shape fill pattern.
    Dotted60Percent,
    /// Dotted 70 percent - shape fill pattern.
    Dotted70Percent,
    /// Dotted 75 percent - shape fill pattern.
    Dotted75Percent,
    /// Dotted 80 percent - shape fill pattern.
    Dotted80Percent,
    /// Dotted 90 percent - shape fill pattern.
    Dotted90Percent,
    /// Diagonal stripes light downwards - shape fill pattern.
    DiagonalStripesLightDownwards,
    /// Diagonal stripes light upwards - shape fill pattern.
    DiagonalStripesLightUpwards,
    /// Diagonal stripes dark downwards - shape fill pattern.
    DiagonalStripesDarkDownwards,
    /// Diagonal stripes dark upwards - shape fill pattern.
    DiagonalStripesDarkUpwards,
    /// Diagonal stripes wide downwards - shape fill pattern.
    DiagonalStripesWideDownwards,
    /// Diagonal stripes wide upwards - shape fill pattern.
    DiagonalStripesWideUpwards,
    /// Vertical stripes light - shape fill pattern.
    VerticalStripesLight,
    /// Horizontal stripes light - shape fill pattern.
    HorizontalStripesLight,
    /// Vertical stripes narrow - shape fill pattern.
    VerticalStripesNarrow,
    /// Horizontal stripes narrow - shape fill pattern.
    HorizontalStripesNarrow,
    /// Vertical stripes dark - shape fill pattern.
    VerticalStripesDark,
    /// Horizontal stripes dark - shape fill pattern.
    HorizontalStripesDark,
    /// Stripes backslashes - shape fill pattern.
    StripesBackslashes,
    /// Stripes forward slashes - shape fill pattern.
    StripesForwardSlashes,
    /// Horizontal stripes alternating - shape fill pattern.
    HorizontalStripesAlternating,
    /// Vertical stripes alternating - shape fill pattern.
    VerticalStripesAlternating,
    /// Small confetti - shape fill pattern.
    SmallConfetti,
    /// Large confetti - shape fill pattern.
    LargeConfetti,
    /// Zigzag - shape fill pattern.
    Zigzag,
    /// Wave - shape fill pattern.
    Wave,
    /// Diagonal brick - shape fill pattern.
    DiagonalBrick,
    /// Horizontal brick - shape fill pattern.
    HorizontalBrick,
    /// Weave - shape fill pattern.
    Weave,
    /// Plaid - shape fill pattern.
    Plaid,
    /// Divot - shape fill pattern.
    Divot,
    /// Dotted grid - shape fill pattern.
    DottedGrid,
    /// Dotted diamond - shape fill pattern.
    DottedDiamond,
    /// Shingle - shape fill pattern.
    Shingle,
    /// Trellis - shape fill pattern.
    Trellis,
    /// Sphere - shape fill pattern.
    Sphere,
    /// Small grid - shape fill pattern.
    SmallGrid,
    /// Large grid - shape fill pattern.
    LargeGrid,
    /// Small checkerboard - shape fill pattern.
    SmallCheckerboard,
    /// Large checkerboard - shape fill pattern.
    LargeCheckerboard,
    /// Outlined diamond grid - shape fill pattern.
    OutlinedDiamondGrid,
    /// Solid diamond grid - shape fill pattern.
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
