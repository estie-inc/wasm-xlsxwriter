use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ConditionalFormatIconTypes` enum defines the conditional
/// format icon types for {@link ConditionalFormatIconSet}.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ConditionalFormatIconType {
    /// Three arrows showing up, sideways and down.
    ThreeArrows,
    /// Three gray arrows showing up, sideways and down.
    ThreeArrowsGray,
    /// Three flags in red, yellow and green.
    ThreeFlags,
    /// Three traffic lights - rounded.
    ThreeTrafficLights,
    /// Three traffic lights with a square rim.
    ThreeTrafficLightsWithRim,
    /// Three shapes like traffic signs - a circle, triangle and diamond.
    ThreeSigns,
    /// Three circled symbols with tick mark, exclamation mark and cross.
    ThreeSymbolsCircled,
    /// Three symbols with tick mark, exclamation mark and cross.
    ThreeSymbols,
    /// Three stars showing different levels of rating.
    ThreeStars,
    /// Three triangles.
    ThreeTriangles,
    /// Four arrows showing up, diagonal up, diagonal down and down.
    FourArrows,
    /// Four gray arrows showing up, diagonal up, diagonal down and down.
    FourArrowsGray,
    /// Four circles in colors going from red to black.
    FourRedToBlack,
    /// Four histogram ratings.
    FourHistograms,
    /// Four traffic lights.
    FourTrafficLights,
    /// Five arrows showing up, diagonal up, sideways, diagonal down and down.
    FiveArrows,
    /// Five gray arrows showing up, diagonal up, sideways, diagonal down and down.
    FiveArrowsGray,
    /// Five histogram ratings.
    FiveHistograms,
    /// Five quarters, from 0 to 4 quadrants filled.
    FiveQuadrants,
    /// Five boxes rating
    FiveBoxes,
}

impl From<ConditionalFormatIconType> for xlsx::ConditionalFormatIconType {
    fn from(value: ConditionalFormatIconType) -> xlsx::ConditionalFormatIconType {
        match value {
            ConditionalFormatIconType::ThreeArrows => xlsx::ConditionalFormatIconType::ThreeArrows,
            ConditionalFormatIconType::ThreeArrowsGray => {
                xlsx::ConditionalFormatIconType::ThreeArrowsGray
            }
            ConditionalFormatIconType::ThreeFlags => xlsx::ConditionalFormatIconType::ThreeFlags,
            ConditionalFormatIconType::ThreeTrafficLights => {
                xlsx::ConditionalFormatIconType::ThreeTrafficLights
            }
            ConditionalFormatIconType::ThreeTrafficLightsWithRim => {
                xlsx::ConditionalFormatIconType::ThreeTrafficLightsWithRim
            }
            ConditionalFormatIconType::ThreeSigns => xlsx::ConditionalFormatIconType::ThreeSigns,
            ConditionalFormatIconType::ThreeSymbolsCircled => {
                xlsx::ConditionalFormatIconType::ThreeSymbolsCircled
            }
            ConditionalFormatIconType::ThreeSymbols => {
                xlsx::ConditionalFormatIconType::ThreeSymbols
            }
            ConditionalFormatIconType::ThreeStars => xlsx::ConditionalFormatIconType::ThreeStars,
            ConditionalFormatIconType::ThreeTriangles => {
                xlsx::ConditionalFormatIconType::ThreeTriangles
            }
            ConditionalFormatIconType::FourArrows => xlsx::ConditionalFormatIconType::FourArrows,
            ConditionalFormatIconType::FourArrowsGray => {
                xlsx::ConditionalFormatIconType::FourArrowsGray
            }
            ConditionalFormatIconType::FourRedToBlack => {
                xlsx::ConditionalFormatIconType::FourRedToBlack
            }
            ConditionalFormatIconType::FourHistograms => {
                xlsx::ConditionalFormatIconType::FourHistograms
            }
            ConditionalFormatIconType::FourTrafficLights => {
                xlsx::ConditionalFormatIconType::FourTrafficLights
            }
            ConditionalFormatIconType::FiveArrows => xlsx::ConditionalFormatIconType::FiveArrows,
            ConditionalFormatIconType::FiveArrowsGray => {
                xlsx::ConditionalFormatIconType::FiveArrowsGray
            }
            ConditionalFormatIconType::FiveHistograms => {
                xlsx::ConditionalFormatIconType::FiveHistograms
            }
            ConditionalFormatIconType::FiveQuadrants => {
                xlsx::ConditionalFormatIconType::FiveQuadrants
            }
            ConditionalFormatIconType::FiveBoxes => xlsx::ConditionalFormatIconType::FiveBoxes,
        }
    }
}
