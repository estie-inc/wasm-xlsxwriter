use std::sync::{Arc, Mutex};

use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

use crate::wrapper::WasmResult;
use crate::macros::wrap_struct;

/// Since the xlsx::Image does not have a default value, we use the smallest PNG image data as a dummy data.
fn new_dummy_image() -> xlsx::Image {
    // Smallest PNG in bytes.
    // https://evanhahn.com/worlds-smallest-png/
    const SMALLEST_PNG_DATA: [u8; 67] = [
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44,
        0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x01, 0x00, 0x00, 0x00, 0x00, 0x37,
        0x6E, 0xF9, 0x24, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41, 0x54, 0x78, 0x01, 0x63, 0x60,
        0x00, 0x00, 0x00, 0x02, 0x00, 0x01, 0x73, 0x75, 0x01, 0x18, 0x00, 0x00, 0x00, 0x00, 0x49,
        0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
    ];
    xlsx::Image::new_from_buffer(&SMALLEST_PNG_DATA).unwrap()
}

/// The `Image` struct is used to create an object to represent an image that
/// can be inserted into a worksheet.

wrap_struct!(
    Image,
    xlsx::Image,
    set_scale_width(scale: f64),
    set_scale_height(scale: f64)
);

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
        let image = xlsx::Image::new_from_buffer(&buffer)?;
        Ok(Image {
            inner: Arc::new(Mutex::new(image)),
        })
    }
}
