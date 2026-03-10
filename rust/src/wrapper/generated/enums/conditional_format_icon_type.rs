use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ConditionalFormatIconTypes` enum defines the conditional
/// format icon types for ConditionalFormatIconSet.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ConditionalFormatIconType {
    /// Three arrows showing up, sideways and down.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/icons_three_arrows.png">
    ThreeArrows,
    /// Three gray arrows showing up, sideways and down.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/icons_three_arrows_gray.png">
    ThreeArrowsGray,
    /// Three flags in red, yellow and green.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/icons_three_flags.png">
    ThreeFlags,
    /// Three traffic lights - rounded.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/icons_three_traffic_lights.png">
    ThreeTrafficLights,
    /// Three traffic lights with a square rim.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/icons_three_traffic_lights_with_rim.png">
    ThreeTrafficLightsWithRim,
    /// Three shapes like traffic signs - a circle, triangle and diamond.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/icons_three_signs.png">
    ThreeSigns,
    /// Three circled symbols with tick mark, exclamation mark and cross.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/icons_three_symbols_circled.png">
    ThreeSymbolsCircled,
    /// Three symbols with tick mark, exclamation mark and cross.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/icons_three_symbols.png">
    ThreeSymbols,
    /// Three stars showing different levels of rating.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/icons_three_stars.png">
    ThreeStars,
    /// Three triangles.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/icons_three_triangles.png">
    ThreeTriangles,
    /// Four arrows showing up, diagonal up, diagonal down and down.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/icons_four_arrows.png">
    FourArrows,
    /// Four gray arrows showing up, diagonal up, diagonal down and down.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/icons_four_arrows_gray.png">
    FourArrowsGray,
    /// Four circles in colors going from red to black.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/icons_four_red_to_black.png">
    FourRedToBlack,
    /// Four histogram ratings.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/icons_four_histograms.png">
    FourHistograms,
    /// Four traffic lights.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/icons_four_traffic_lights.png">
    FourTrafficLights,
    /// Five arrows showing up, diagonal up, sideways, diagonal down and down.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/icons_five_arrows.png">
    FiveArrows,
    /// Five gray arrows showing up, diagonal up, sideways, diagonal down and down.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/icons_five_arrows_gray.png">
    FiveArrowsGray,
    /// Five histogram ratings.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/icons_five_histograms.png">
    FiveHistograms,
    /// Five quarters, from 0 to 4 quadrants filled.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/icons_five_quadrants.png">
    FiveQuadrants,
    /// Five boxes rating
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/icons_five_boxes.png">
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
