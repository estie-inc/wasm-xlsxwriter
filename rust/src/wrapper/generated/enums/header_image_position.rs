use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `HeaderImagePosition` enum defines the image position in a header or footer.
///
/// Used with the
/// Worksheet.setHeaderImage()(crate::Worksheet::set_header_image) and
/// Worksheet.setFooterImage()(crate::Worksheet::set_footer_image)
/// methods.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum HeaderImagePosition {
    /// The image is positioned in the left section of the header/footer.
    Left,
    /// The image is positioned in the center section of the header/footer.
    Center,
    /// The image is positioned in the right section of the header/footer.
    Right,
}

impl From<HeaderImagePosition> for xlsx::HeaderImagePosition {
    fn from(value: HeaderImagePosition) -> xlsx::HeaderImagePosition {
        match value {
            HeaderImagePosition::Left => xlsx::HeaderImagePosition::Left,
            HeaderImagePosition::Center => xlsx::HeaderImagePosition::Center,
            HeaderImagePosition::Right => xlsx::HeaderImagePosition::Right,
        }
    }
}
