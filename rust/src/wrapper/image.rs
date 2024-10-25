use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

use crate::wrapper::WasmResult;

use super::map_xlsx_error;

/// The `Image` struct is used to create an object to represent an image that
/// can be inserted into a worksheet.
#[derive(Clone)]
#[wasm_bindgen]
pub struct Image {
    pub(crate) inner: xlsx::Image,
}

#[wasm_bindgen]
impl Image {
    /// Create an Image object from a u8 buffer. The image can then be inserted
    /// into a worksheet.
    ///
    /// The supported image formats are:
    ///
    /// - PNG
    /// - JPG
    /// - GIF: The image can be an animated gif in more recent versions of
    ///   Excel.
    /// - BMP: BMP images are only supported for backward compatibility. In
    ///   general it is best to avoid BMP images since they are not compressed.
    ///   If used, BMP images must be 24 bit, true color, bitmaps.
    ///
    /// EMF and WMF file formats will be supported in an upcoming version of the
    /// library.
    ///
    /// **NOTE on SVG files**: Excel doesn't directly support SVG files in the
    /// same way as other image file formats. It allows SVG to be inserted into
    /// a worksheet but converts them to, and displays them as, PNG files. It
    /// stores the original SVG image in the file so the original format can be
    /// retrieved. This removes the file size and resolution advantage of using
    /// SVG files. As such SVG files are not supported by `rust_xlsxwriter`
    /// since a conversion to the PNG format would be required and that format
    /// is already supported.
    ///
    /// @param {array} buffer - The image data as a u8 array or vector.
    /// @returns {Image} - The Image object.
    ///
    /// TODO: error
    /// - [`XlsxError::UnknownImageType`] - Unknown image type. The supported
    ///   image formats are PNG, JPG, GIF and BMP.
    /// - [`XlsxError::ImageDimensionError`] - Image has 0 width or height, or
    ///   the dimensions couldn't be read.
    #[wasm_bindgen(constructor, skip_jsdoc)]
    pub fn new(buffer: Vec<u8>) -> WasmResult<Image> {
        let image = map_xlsx_error(xlsx::Image::new_from_buffer(&buffer))?;
        Ok(Image { inner: image })
    }

    /// Set the width scale for the image.
    ///
    /// Set the width scale for the image relative to 1.0 (i.e. 100%). See the
    /// {@link Image#setScaleHeight} method for details.
    ///
    /// @param {number} scale - The scale ratio.
    /// @returns {Image} - The Image object.
    #[wasm_bindgen(js_name = "setScaleWidth", skip_jsdoc)]
    pub fn set_scale_width(&self, scale: f64) -> Image {
        Image {
            inner: self.clone().inner.set_scale_width(scale),
        }
    }

    /// Set the width scale for the image.
    ///
    /// Set the width scale for the image relative to 1.0 (i.e. 100%). See the
    /// {@link Image#setScaleHeight} method for details.
    ///
    /// @param {number} scale - The scale ratio.
    /// @returns {Image} - The Image object.
    #[wasm_bindgen(js_name = "setScaleHeight", skip_jsdoc)]
    pub fn set_scale_height(&self, scale: f64) -> Image {
        Image {
            inner: self.clone().inner.set_scale_height(scale),
        }
    }
}
