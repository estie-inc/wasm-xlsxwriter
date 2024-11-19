use std::sync::{Arc, Mutex, MutexGuard};

use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

use super::color::Color;

/// The `Format` struct is used to define cell formatting for data in a
/// worksheet.
///
/// The properties of a cell that can be formatted include: fonts, colors,
/// patterns, borders, alignment and number formatting.
///
/// TODO: example omitted
///
/// # Creating and using a Format object
///
/// Formats are created by calling the {@link Format::constructor} method and properties as
/// set using the various methods shown is this section of the document. Once
/// the Format has been created it can be passed to one of the worksheet
/// `write*()` methods. Multiple properties can be set by chaining them
/// together:
///
/// TODO: example omitted
///
/// # Format methods and Format properties
///
/// The following table shows the Excel format categories, in the order shown in
/// the Excel "Format Cell" dialog, and the equivalent `rust_xlsxwriter` Format
/// method:
///
/// | Category        | Description           |  Method Name                             |
/// | :-------------- | :-------------------- |  :-------------------------------------- |
/// | **Number**      | Numeric format        |  {@link Format#setNumFormat}            |
/// | **Alignment**   | Horizontal align      |  {@link Format#setAlign}                 |
/// |                 | Vertical align        |  {@link Format#setAlign}                 |
/// |                 | Rotation              |  {@link Format#setRotation}              |
/// |                 | Text wrap             |  {@link Format#setTextWrap}             |
/// |                 | Indentation           |  {@link Format#setIndent}                |
/// |                 | Reading direction     |  {@link Format#setReadingDirection}     |
/// |                 | Shrink to fit         |  {@link Format#setShrink}                |
/// | **Font**        | Font type             |  {@link Format#setFontName}             |
/// |                 | Font size             |  {@link Format#setFontSize}             |
/// |                 | Font color            |  {@link Format#setFontColor}            |
/// |                 | Bold                  |  {@link Format#setBold}                  |
/// |                 | Italic                |  {@link Format#setItalic}                |
/// |                 | Underline             |  {@link Format#setUnderline}             |
/// |                 | Strikethrough         |  {@link Format#setFontStrikethrough}    |
/// |                 | Super/Subscript       |  {@link Format#setFontScript}           |
/// | **Border**      | Cell border           |  {@link Format#setBorder}                |
/// |                 | Bottom border         |  {@link Format#setBorderBottom}         |
/// |                 | Top border            |  {@link Format#setBorderTop}            |
/// |                 | Left border           |  {@link Format#setBorderLeft}           |
/// |                 | Right border          |  {@link Format#setBorderRight}          |
/// |                 | Border color          |  {@link Format#setBorderColor}          |
/// |                 | Bottom color          |  {@link Format#setBorderBottomColor}   |
/// |                 | Top color             |  {@link Format#setBorderTopColor}      |
/// |                 | Left color            |  {@link Format#setBorderLeftColor}     |
/// |                 | Right color           |  {@link Format#setBorderRightColor}    |
/// |                 | Diagonal border       |  {@link Format#setBorderDiagonal}       |
/// |                 | Diagonal border color |  {@link Format#setBorderDiagonalColor} |
/// |                 | Diagonal border type  |  {@link Format#setBorderDiagonalType}  |
/// | **Fill**        | Cell pattern          |  {@link Format#setPattern}               |
/// |                 | Background color      |  {@link Format#setBackgroundColor}      |
/// |                 | Foreground color      |  {@link Format#setForegroundColor}      |
/// | **Protection**  | Unlock cells          |  {@link Format#setUnlocked}              |
/// |                 | Hide formulas         |  {@link Format#setHidden}                |
///
/// # Format Colors
///
/// Format property colors are specified by using the {@link Color} enum with a Html
/// style RGB integer value or a limited number of defined colors:
///
/// TODO: example omitted
///
/// # Format Defaults
///
/// The default Excel 365 cell format is a font setting of Calibri size 11 with
/// all other properties turned off.
///
/// It is occasionally useful to use a default format with a method that
/// requires a format but where you don't actually want to change the
/// formatting.
///
/// TODO: example omitted
///
/// # Number Format Categories
///
/// The {@link Format::setNumFormat} method is used to set the number format for
/// numbers used with
/// {@link Worksheet#writeNumberWithFormat}(crate::Worksheet::write_number_with_format()):
///
/// TODO: example omitted
///
/// If the number format you use is the same as one of Excel's built in number
/// formats then it will have a number category other than "General" or
/// "Number". The Excel number categories are:
///
/// - General
/// - Number
/// - Currency
/// - Accounting
/// - Date
/// - Time
/// - Percentage
/// - Fraction
/// - Scientific
/// - Text
/// - Custom
///
/// In the case of the example above the formatted output shows up as a Number
/// category:
///
/// <img src="https://rustxlsxwriter.github.io/images/format_currency1.png">
///
/// If we wanted to have the number format display as a different category, such
/// as Currency, then would need to match the number format string used in the
/// code with the number format used by Excel. The easiest way to do this is to
/// open the Number Formatting dialog in Excel and set the required format:
///
/// <img src="https://rustxlsxwriter.github.io/images/format_currency2.png">
///
/// Then, while still in the dialog, change to Custom. The format displayed is
/// the format used by Excel.
///
/// <img src="https://rustxlsxwriter.github.io/images/format_currency3.png">
///
/// If we put the format that we found (`"[$$-409]#,##0.00"`) into our previous
/// example and rerun it we will get a number format in the Currency category:
///
/// TODO: example omitted
///
/// That give us the following updated output. Note that the number category is
/// now shown as Currency:
///
/// <img src="https://rustxlsxwriter.github.io/images/format_currency4.png">
///
/// The same process can be used to find format strings for "Date" or
/// "Accountancy" formats.
///
/// # Number Formats in different locales
///
/// As shown in the previous section the `format.set_num_format()` method is
/// used to set the number format for `rust_xlsxwriter` formats. A common use
/// case is to set a number format with a "grouping/thousands" separator and a
/// "decimal" point:
///
/// TODO: example omitted
///
/// In the US locale (and some others) where the number "grouping/thousands"
/// separator is `","` and the "decimal" point is `"."` which would be shown in
/// Excel as:
///
/// <img src="https://rustxlsxwriter.github.io/images/format_currency5.png">
///
/// In other locales these values may be reversed or different. They are
/// generally set in the "Region" settings of Windows or Mac OS.  Excel handles
/// this by storing the number format in the file format in the US locale, in
/// this case `#,##0.00`, but renders it according to the regional settings of
/// the host OS. For example, here is the same, unmodified, output file shown
/// above in a German locale:
///
/// <img src="https://rustxlsxwriter.github.io/images/format_currency6.png">
///
/// And here is the same file in a Russian locale. Note the use of a space as
/// the "grouping/thousands" separator:
///
/// <img src="https://rustxlsxwriter.github.io/images/format_currency7.png">
///
/// In order to replicate Excel's behavior all `rust_xlsxwriter` programs should
/// use US locale formatting which will then be rendered in the settings of your
/// host OS.
#[derive(Clone)]
#[wasm_bindgen]
pub struct Format {
    pub(crate) inner: Arc<Mutex<xlsx::Format>>,
}

macro_rules! impl_method {
    ($self:ident.$method:ident($($arg:expr),*)) => {
        let mut lock = $self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.$method($($arg),*);
        let _ = std::mem::replace(&mut *lock, inner);
        return Format {
            inner: Arc::clone(&$self.inner),
        }
    };
}

#[wasm_bindgen]
impl Format {
    /// Create a new Format object.
    ///
    /// Create a new Format object to use with worksheet formatting.
    #[wasm_bindgen(constructor)]
    pub fn new() -> Format {
        Format {
            inner: Arc::new(Mutex::new(xlsx::Format::new())),
        }
    }

    pub(crate) fn lock(&self) -> MutexGuard<xlsx::Format> {
        self.inner.lock().unwrap()
    }

    /// Set the Format alignment properties.
    ///
    /// This method is used to set the horizontal and vertical data alignment
    /// within a cell.
    ///
    /// @param {FormatAlign} align - The vertical and or horizontal alignment direction as defined by the {@link FormatAlign} enum.
    /// @return {Format} - The Format instance.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "setAlign", skip_jsdoc)]
    pub fn set_align(&self, align: FormatAlign) -> Format {
        impl_method!(self.set_align(align.into()));
    }

    /// Set the bold property for a Format font.
    ///
    /// @return {Format} - The Format instance.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "setBold", skip_jsdoc)]
    pub fn set_bold(&self) -> Format {
        impl_method!(self.set_bold());
    }

    /// Set the italic property for the Format font.
    ///
    /// @return {Format} - The Format instance.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "setItalic", skip_jsdoc)]
    pub fn set_italic(&self) -> Format {
        impl_method!(self.set_italic());
    }

    #[wasm_bindgen(js_name = "setUnderline")]
    pub fn set_underline(&self, underline: FormatUnderline) -> Format {
        impl_method!(self.set_underline(underline.into()));
    }

    #[wasm_bindgen(js_name = "setTextWrap")]
    pub fn set_text_wrap(&self) -> Format {
        impl_method!(self.set_text_wrap());
    }

    #[wasm_bindgen(js_name = "setIndent")]
    pub fn set_indent(&self, level: u8) -> Format {
        impl_method!(self.set_indent(level));
    }

    /// Set the Format rotation property.
    ///
    /// Set the rotation angle of the text in a cell. The rotation can be any
    /// angle in the range -90 to 90 degrees, or 270 to indicate text where the
    /// letters run from top to bottom.
    ///
    /// @param {number} rotation - The rotation angle.
    /// @return {Format} - The Format instance.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "setRotation", skip_jsdoc)]
    pub fn set_rotation(&self, rotation: i16) -> Format {
        impl_method!(self.set_rotation(rotation));
    }

    /// Set the Format border property.
    ///
    /// Set the cell border style. Individual border elements can be configured
    /// using the following methods with the same parameters:
    ///
    /// - {@link Format#setBorderTop}
    /// - {@link Format#setBorderLeft}
    /// - {@link Format#setBorderRight}
    /// - {@link Format#setBorderBottom}
    ///
    /// @param {FormatBorder} border - The border property as defined by a {@link FormatBorder} enum
    ///   value.
    /// @return {Format} - The Format instance.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "setBorder", skip_jsdoc)]
    pub fn set_border(&self, border: FormatBorder) -> Format {
        impl_method!(self.set_border(border.into()));
    }

    /// Set the Format border color property.
    ///
    /// Set the cell border color. Individual border elements can be configured
    /// using the following methods with the same parameters:
    ///
    /// - {@link Format#setBorderTopColor}
    /// - {@link Format#setBorderLeftColor}
    /// - {@link Format#setBorderRightColor}
    /// - {@link Format#setBorderBottomColor}
    ///
    /// @param {Color} color - The border color as defined by a {@link Color}.
    /// @return {Format} - The Format instance.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "setBorderColor", skip_jsdoc)]
    pub fn set_border_color(&self, color: Color) -> Format {
        impl_method!(self.set_border_color(color.inner));
    }

    /// Set the cell bottom border style.
    ///
    /// See {@link Format#setBorder} for details.
    ///
    /// @param {FormatBorder} border - The border property as defined by a {@link FormatBorder} enum
    ///   value.
    /// @return {Format} - The Format instance.
    #[wasm_bindgen(js_name = "setBorderBottom", skip_jsdoc)]
    pub fn set_border_bottom(&self, border: FormatBorder) -> Format {
        impl_method!(self.set_border_bottom(border.into()));
    }

    /// Set the cell bottom border color.
    ///
    /// See {@link Format#setBorderColor} for details.
    ///
    /// @param {Color} color - The border color as defined by a {@link Color}.
    /// @return {Format} - The Format instance.
    #[wasm_bindgen(js_name = "setBorderBottomColor", skip_jsdoc)]
    pub fn set_border_bottom_color(&self, color: Color) -> Format {
        impl_method!(self.set_border_bottom_color(color.inner));
    }

    /// Set the cell top border style.
    ///
    /// See {@link Format#setBorder} for details.
    ///
    /// @param {FormatBorder} border - The border property as defined by a {@link FormatBorder} enum
    ///   value.
    /// @return {Format} - The Format instance.
    #[wasm_bindgen(js_name = "setBorderTop", skip_jsdoc)]
    pub fn set_border_top(&self, border: FormatBorder) -> Format {
        impl_method!(self.set_border_top(border.into()));
    }

    /// Set the cell top border color.
    ///
    /// See {@link Format#setBorderColor} for details.
    ///
    /// @param {Color} color - The border color as defined by a {@link Color}.
    /// @return {Format} - The Format instance.
    #[wasm_bindgen(js_name = "setBorderTopColor", skip_jsdoc)]
    pub fn set_border_top_color(&self, color: Color) -> Format {
        impl_method!(self.set_border_top_color(color.inner));
    }

    /// Set the cell left border style.
    ///
    /// See {@link Format#setBorder} for details.
    ///
    /// @param {FormatBorder} border - The border property as defined by a {@link FormatBorder} enum
    ///   value.
    /// @return {Format} - The Format instance.
    #[wasm_bindgen(js_name = "setBorderLeft", skip_jsdoc)]
    pub fn set_border_left(&self, border: FormatBorder) -> Format {
        impl_method!(self.set_border_left(border.into()));
    }

    /// Set the cell left border color.
    ///
    /// See {@link Format#setBorderColor} for details.
    ///
    /// @param {Color} color - The border color as defined by a {@link Color}.
    /// @return {Format} - The Format instance.
    #[wasm_bindgen(js_name = "setBorderLeftColor", skip_jsdoc)]
    pub fn set_border_left_color(&self, color: Color) -> Format {
        impl_method!(self.set_border_left_color(color.inner));
    }

    /// Set the cell right border style.
    ///
    /// See {@link Format#setBorder} for details.
    ///
    /// @param {FormatBorder} border - The border property as defined by a {@link FormatBorder} enum
    ///   value.
    /// @return {Format} - The Format instance.
    #[wasm_bindgen(js_name = "setBorderRight", skip_jsdoc)]
    pub fn set_border_right(&self, border: FormatBorder) -> Format {
        impl_method!(self.set_border_right(border.into()));
    }

    /// Set the cell right border color.
    ///
    /// See {@link Format#setBorderColor} for details.
    ///
    /// @param {Color} color - The border color as defined by a {@link Color}.
    /// @return {Format} - The Format instance.
    #[wasm_bindgen(js_name = "setBorderRightColor", skip_jsdoc)]
    pub fn set_border_right_color(&self, color: Color) -> Format {
        impl_method!(self.set_border_right_color(color.inner));
    }

    /// Set the Format border diagonal property.
    ///
    /// Set the cell border diagonal line style. This method should be used in
    /// conjunction with the {@link Format#setBorderDiagonalType} method to
    /// set the diagonal type.
    ///
    /// @param {FormatBorder} border - The border property as defined by a {@link FormatBorder} enum
    ///   value.
    /// @return {Format} - The Format instance.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "setBorderDiagonal", skip_jsdoc)]
    pub fn set_border_diagonal(&self, border: FormatBorder) -> Format {
        impl_method!(self.set_border_diagonal(border.into()));
    }

    /// Set the cell diagonal border color.
    ///
    /// See {@link Format#setBorderDiagonal} for details.
    ///
    /// @param {Color} color - The border color as defined by a {@link Color}.
    /// @return {Format} - The Format instance.
    #[wasm_bindgen(js_name = "setBorderDiagonalColor", skip_jsdoc)]
    pub fn set_border_diagonal_color(&self, color: Color) -> Format {
        impl_method!(self.set_border_diagonal_color(color.inner));
    }

    /// Set the color property for the Format font.
    ///
    /// The `setFontColor()` method is used to set the font color in a cell.
    /// To set the color of a cell background use the `setBackgroundColor()` and
    /// `setPattern()` methods.
    ///
    /// @param {Color} color - The font color property defined by a {@link Color} enum
    ///   value.
    /// @return {Format} - The Format instance.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "setFontColor", skip_jsdoc)]
    pub fn set_font_color(&self, color: Color) -> Format {
        impl_method!(self.set_font_color(color.inner));
    }

    /// Set the Format font family property.
    ///
    /// Set the font family. This is usually an integer in the range 1-4. This
    /// function is implemented for completeness but is rarely used in practice.
    ///
    /// @param {number} font_family - The font family property.
    /// @return {Format} - The Format instance.
    #[wasm_bindgen(js_name = "setFontFamily", skip_jsdoc)]
    pub fn set_font_family(&self, font_family: u8) -> Format {
        impl_method!(self.set_font_family(font_family));
    }

    /// Set the Format font name property.
    ///
    /// Set the font for a cell format. Excel can only display fonts that are
    /// installed on the system that it is running on. Therefore it is generally
    /// best to use standard Excel fonts.
    ///
    /// @param {string} font_name - The font name property.
    /// @return {Format} - The Format instance.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "setFontName")]
    pub fn set_font_name(&self, font_name: &str) -> Format {
        impl_method!(self.set_font_name(font_name));
    }

    /// Set the Format font size property.
    ///
    /// Set the font size of the cell format. The size is generally an integer
    /// value but Excel allows x.5 values (hence the property is converted into a f64).
    ///
    /// Excel adjusts the height of a row to accommodate the largest font size
    /// in the row.
    ///
    /// @param {number} font_size - The font size property.
    /// @return {Format} - The Format instance.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "setFontSize", skip_jsdoc)]
    pub fn set_font_size(&self, font_size: f64) -> Format {
        impl_method!(self.set_font_size(font_size));
    }

    /// Set the Format font scheme property.
    ///
    /// This function is implemented for completeness but is rarely used in
    /// practice.
    ///
    /// @param {string} font_scheme - The font scheme property.
    /// @return {Format} - The Format instance.
    #[wasm_bindgen(js_name = "setFontScheme", skip_jsdoc)]
    pub fn set_font_scheme(&self, font_scheme: &str) -> Format {
        impl_method!(self.set_font_scheme(font_scheme));
    }

    /// Set the Format font character set property.
    ///
    /// Set the font character. This function is implemented for completeness
    /// but is rarely used in practice.
    ///
    /// @param {number} font_charset - The font character set property.
    /// @return {Format} - The Format instance.
    #[wasm_bindgen(js_name = "setFontCharset", skip_jsdoc)]
    pub fn set_font_charset(&self, font_charset: u8) -> Format {
        impl_method!(self.set_font_charset(font_charset));
    }

    /// Set the Format font strikethrough property.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "setFontStrikethrough")]
    pub fn set_font_strikethrough(&self) -> Format {
        impl_method!(self.set_font_strikethrough());
    }

    /// Set the Format pattern foreground color property.
    ///
    /// The `set_foreground_color` method can be used to set the
    /// foreground/pattern color of a pattern. Patterns are defined via the
    /// {@link Format#setPattern} method.
    ///
    /// @param {Color} color - The foreground color property defined by a {@link Color}.
    /// @return {Format} - The Format instance.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "setForegroundColor", skip_jsdoc)]
    pub fn set_foreground_color(&self, color: Color) -> Format {
        impl_method!(self.set_foreground_color(color.inner));
    }

    /// Set the Format pattern background color property.
    ///
    /// The `setBackgroundColor` method can be used to set the background
    /// color of a pattern. Patterns are defined via the {@link Format#setPattern}
    /// method. If a pattern hasn't been defined then a solid fill pattern is
    /// used as the default.
    ///
    /// @param {Color} color - The background color property defined by a {@link Color}.
    /// @return {Format} - The Format instance.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "setBackgroundColor", skip_jsdoc)]
    pub fn set_background_color(&self, color: Color) -> Format {
        impl_method!(self.set_background_color(color.inner));
    }

    /// Set the number format for a Format.
    ///
    /// This method is used to define the numerical format of a number in Excel.
    /// It controls whether a number is displayed as an integer, a floating
    /// point number, a date, a currency value or some other user defined
    /// format.
    ///
    /// See also [Number Format Categories] and [Number Formats in different
    /// locales].
    ///
    /// [Number Format Categories]: crate::Format#number-format-categories
    /// [Number Formats in different locales]:
    ///     crate::Format#number-formats-in-different-locales
    ///
    /// @param {string} num_format - The number format property.
    /// @return {Format} - The Format instance.
    #[wasm_bindgen(js_name = "setNumFormat", skip_jsdoc)]
    pub fn set_num_format(&self, num_format: &str) -> Format {
        impl_method!(self.set_num_format(num_format));
    }

    /// Set the Format pattern property.
    ///
    /// Set the pattern for a cell. The most commonly used pattern is
    /// {@link FormatPattern#Solid}.
    ///
    /// To set the pattern colors see {@link Format#setBackgroundColor} and
    /// {@link Format#setForegroundColor}
    ///
    /// @param {FormatPattern} pattern - The pattern property defined by a {@link FormatPattern} enum
    ///   value.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "setPattern", skip_jsdoc)]
    pub fn set_pattern(&self, pattern: FormatPattern) -> Format {
        impl_method!(self.set_pattern(pattern.into()));
    }

    /// Set the Format property to hide formulas in a cell.
    ///
    /// This method can be used to hide a formula while still displaying its
    /// result. This is generally used to hide complex calculations from end
    /// users who are only interested in the result. It only has an effect if
    /// the worksheet has been protected using the
    /// {@link Worksheet#protect} method.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "setHidden", skip_jsdoc)]
    pub fn set_hidden(&self) -> Format {
        impl_method!(self.set_hidden());
    }

    #[wasm_bindgen(js_name = "setLocked")]
    pub fn set_locked(&self) -> Format {
        impl_method!(self.set_locked());
    }

    #[wasm_bindgen(js_name = "setUnlocked")]
    pub fn set_unlocked(&self) -> Format {
        impl_method!(self.set_unlocked());
    }
}

/// The `FormatAlign` enum defines the vertical and horizontal alignment properties
/// of a {@link Format}.
#[derive(Debug, Clone, Copy)]
#[wasm_bindgen]
pub enum FormatAlign {
    /// General/default alignment. The cell will use Excel's default for the
    /// data type, for example Left for text and Right for numbers.
    General,
    /// Align text to the left.
    Left,
    /// Center text horizontally.
    Center,
    /// Align text to the right.
    Right,
    /// Fill (repeat) the text horizontally across the cell.
    Fill,
    /// Aligns the text to the left and right of the cell, if the text exceeds
    /// the width of the cell.
    Justify,
    /// Center the text across the cell or cells that have this alignment. This
    /// is an older form of merged cells.
    CenterAcross,
    /// Distribute the words in the text evenly across the cell.
    Distributed,
    /// Align text to the top.
    Top,
    /// Align text to the bottom.
    Bottom,
    /// Center text vertically.
    VerticalCenter,
    /// Aligns the text to the top and bottom of the cell, if the text exceeds
    /// the height of the cell.
    VerticalJustify,
    /// Distribute the words in the text evenly from top to bottom in the cell.
    VerticalDistributed,
}

impl From<FormatAlign> for xlsx::FormatAlign {
    fn from(align: FormatAlign) -> xlsx::FormatAlign {
        match align {
            FormatAlign::General => xlsx::FormatAlign::General,
            FormatAlign::Left => xlsx::FormatAlign::Left,
            FormatAlign::Center => xlsx::FormatAlign::Center,
            FormatAlign::Right => xlsx::FormatAlign::Right,
            FormatAlign::Fill => xlsx::FormatAlign::Fill,
            FormatAlign::Justify => xlsx::FormatAlign::Justify,
            FormatAlign::CenterAcross => xlsx::FormatAlign::CenterAcross,
            FormatAlign::Distributed => xlsx::FormatAlign::Distributed,
            FormatAlign::Top => xlsx::FormatAlign::Top,
            FormatAlign::Bottom => xlsx::FormatAlign::Bottom,
            FormatAlign::VerticalCenter => xlsx::FormatAlign::VerticalCenter,
            FormatAlign::VerticalJustify => xlsx::FormatAlign::VerticalJustify,
            FormatAlign::VerticalDistributed => xlsx::FormatAlign::VerticalDistributed,
        }
    }
}

/// The `FormatBorder` enum defines the Excel border types that can be added to
/// a {@link Format} pattern.
#[derive(Debug, Clone, Copy)]
#[wasm_bindgen]
pub enum FormatBorder {
    /// No border.
    None,
    /// Thin border style.
    Thin,
    /// Medium border style.
    Medium,
    /// Dashed border style.
    Dashed,
    /// Dotted border style.
    Dotted,
    /// Thick border style.
    Thick,
    /// Double border style.
    Double,
    /// Hair border style.
    Hair,
    /// Medium dashed border style.
    MediumDashed,
    /// Dash-dot border style.
    DashDot,
    /// Medium dash-dot border style.
    MediumDashDot,
    /// Dash-dot-dot border style.
    DashDotDot,
    /// Medium dash-dot-dot border style.
    MediumDashDotDot,
    /// Slant dash-dot border style.
    SlantDashDot,
}

impl From<FormatBorder> for xlsx::FormatBorder {
    fn from(border: FormatBorder) -> xlsx::FormatBorder {
        match border {
            FormatBorder::None => xlsx::FormatBorder::None,
            FormatBorder::Thin => xlsx::FormatBorder::Thin,
            FormatBorder::Medium => xlsx::FormatBorder::Medium,
            FormatBorder::Dashed => xlsx::FormatBorder::Dashed,
            FormatBorder::Dotted => xlsx::FormatBorder::Dotted,
            FormatBorder::Thick => xlsx::FormatBorder::Thick,
            FormatBorder::Double => xlsx::FormatBorder::Double,
            FormatBorder::Hair => xlsx::FormatBorder::Hair,
            FormatBorder::MediumDashed => xlsx::FormatBorder::MediumDashed,
            FormatBorder::DashDot => xlsx::FormatBorder::DashDot,
            FormatBorder::MediumDashDot => xlsx::FormatBorder::MediumDashDot,
            FormatBorder::DashDotDot => xlsx::FormatBorder::DashDotDot,
            FormatBorder::MediumDashDotDot => xlsx::FormatBorder::MediumDashDotDot,
            FormatBorder::SlantDashDot => xlsx::FormatBorder::SlantDashDot,
        }
    }
}

#[derive(Debug, Clone, Copy)]
#[wasm_bindgen]
pub enum FormatPattern {
    /// Automatic or Empty pattern.
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
    fn from(pattern: FormatPattern) -> xlsx::FormatPattern {
        match pattern {
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

/// The `FormatUnderline` enum defines the font underline type in a {@link Format}.
///
/// The difference between a normal underline and an "accounting" underline is
/// that a normal underline only underlines the text/number in a cell whereas an
/// accounting underline underlines the entire cell width.
///
/// TODO: example omitted
#[derive(Clone, Copy)]
#[wasm_bindgen]
pub enum FormatUnderline {
    /// The default/automatic underline for an Excel font.
    None,
    /// A single underline under the text/number in a cell.
    Single,
    /// A double underline under the text/number in a cell.
    Double,
    /// A single accounting style underline under the entire cell.
    SingleAccounting,
    /// A double accounting style underline under the entire cell.
    DoubleAccounting,
}

impl From<FormatUnderline> for xlsx::FormatUnderline {
    fn from(underline: FormatUnderline) -> xlsx::FormatUnderline {
        match underline {
            FormatUnderline::None => xlsx::FormatUnderline::None,
            FormatUnderline::Single => xlsx::FormatUnderline::Single,
            FormatUnderline::Double => xlsx::FormatUnderline::Double,
            FormatUnderline::SingleAccounting => xlsx::FormatUnderline::SingleAccounting,
            FormatUnderline::DoubleAccounting => xlsx::FormatUnderline::DoubleAccounting,
        }
    }
}
