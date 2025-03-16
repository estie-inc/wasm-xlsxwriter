use std::sync::{Arc, Mutex, MutexGuard};

use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

use super::color::Color;
use crate::macros::wrap_struct;

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
/// |                 | Border color          |  {@link Format#setBorderColor}          |
/// | **Fill**        | Cell pattern          |  {@link Format#setPattern}               |
/// |                 | Background color      |  {@link Format#setBackgroundColor}      |
/// |                 | Foreground color      |  {@link Format#setForegroundColor}      |
/// | **Protection**  | Lock cells            |  {@link Format#setLocked}                |
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
wrap_struct!(
    Format,
    xlsx::Format,
    set_align(align: FormatAlign),
    set_bold(),
    set_italic(),
    set_underline(underline: FormatUnderline),
    set_text_wrap(),
    set_indent(level: u8),
    set_rotation(rotation: i16),
    set_border(border: FormatBorder),
    set_border_color(color: Color),
    set_border_bottom(border: FormatBorder),
    set_border_bottom_color(color: Color),
    set_border_top(border: FormatBorder),
    set_border_top_color(color: Color),
    set_border_left(border: FormatBorder),
    set_border_left_color(color: Color),
    set_border_right(border: FormatBorder),
    set_border_right_color(color: Color),
    set_border_diagonal(border: FormatBorder),
    set_border_diagonal_color(color: Color),
    set_border_diagonal_type(border: FormatDiagonalBorder),
    set_hyperlink(),
    set_font_color(color: Color),
    set_font_family(font_family: u8),
    set_font_name(font_name: &str),
    set_font_size(font_size: f64),
    set_font_scheme(font_scheme: &str),
    set_font_charset(font_charset: u8),
    set_font_strikethrough(),
    set_format_script(script: FormatScript),
    set_foreground_color(color: Color),
    set_background_color(color: Color),
    set_num_format(num_format: &str),
    set_pattern(pattern: FormatPattern),
    set_hidden(),
    set_locked(),
    set_unlocked(),
    set_quote_prefix()
);

#[wasm_bindgen]
impl Format {
    /// Create a deep clone of the Format object.
    ///
    /// Create a deep clone of the Format object. This is useful when you want
    /// to create a new format based on an existing format but don't want the
    /// changes to affect the original format.
    ///
    /// @returns {Format} - A new Format object.
    #[wasm_bindgen(js_name = "clone")]
    pub fn deep_clone(&self) -> Format {
        let inner = self.lock();
        Format {
            inner: Arc::new(Mutex::new(inner.clone())),
        }
    }
}

/// The `FormatAlign` enum defines the alignment options for a cell.
///
/// The `FormatAlign` enum defines the alignment options for a cell. It is used
/// in conjunction with the {@link Format#setAlign} method.
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

/// The `FormatBorder` enum defines the border styles for a cell.
///
/// The `FormatBorder` enum defines the border styles for a cell. It is used in
/// conjunction with the {@link Format#setBorder} method.
#[wasm_bindgen]
#[derive(Default)]
pub enum FormatBorder {
    /// No border.
    #[default]
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

/// The `FormatDiagonalBorder` enum defines the diagonal border styles for a cell.
///
/// The `FormatDiagonalBorder` enum defines the diagonal border styles for a
/// cell. It is used in conjunction with the {@link Format#setBorderDiagonalType}
/// method.
#[wasm_bindgen]
#[derive(Default)]
pub enum FormatDiagonalBorder {
    /// The default/automatic format for an Excel font.
    #[default]
    None,
    /// Cell diagonal border from bottom left to top right.
    BorderUp,
    /// Cell diagonal border from top left to bottom right.
    BorderDown,
    /// Cell diagonal border in both directions.
    BorderUpDown,
}

impl From<FormatDiagonalBorder> for xlsx::FormatDiagonalBorder {
    fn from(border: FormatDiagonalBorder) -> xlsx::FormatDiagonalBorder {
        match border {
            FormatDiagonalBorder::None => xlsx::FormatDiagonalBorder::None,
            FormatDiagonalBorder::BorderUp => xlsx::FormatDiagonalBorder::BorderUp,
            FormatDiagonalBorder::BorderDown => xlsx::FormatDiagonalBorder::BorderDown,
            FormatDiagonalBorder::BorderUpDown => xlsx::FormatDiagonalBorder::BorderUpDown,
        }
    }
}

/// The `FormatPattern` enum defines the pattern styles for a cell.
///
/// The `FormatPattern` enum defines the pattern styles for a cell. It is used
/// in conjunction with the {@link Format#setPattern} method.
#[wasm_bindgen]
#[derive(Default)]
pub enum FormatPattern {
    /// Automatic or Empty pattern.
    #[default]
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

/// The `FormatUnderline` enum defines the underline styles for a cell.
///
/// The `FormatUnderline` enum defines the underline styles for a cell. It is
/// used in conjunction with the {@link Format#setUnderline} method.
#[wasm_bindgen]
#[derive(Default)]
pub enum FormatUnderline {
    /// The default/automatic underline for an Excel font.
    #[default]
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

/// The `FormatScript` enum defines the script styles for a cell.
///
/// The `FormatScript` enum defines the script styles for a cell. It is used in
/// conjunction with the {@link Format#setFormatScript} method.
#[wasm_bindgen]
#[derive(Default)]
pub enum FormatScript {
    /// The default/automatic format for an Excel font.
    #[default]
    None,
    /// The cell text is superscripted.
    Superscript,
    /// The cell text is subscripted.
    Subscript,
}

impl From<FormatScript> for xlsx::FormatScript {
    fn from(script: FormatScript) -> xlsx::FormatScript {
        match script {
            FormatScript::None => xlsx::FormatScript::None,
            FormatScript::Superscript => xlsx::FormatScript::Superscript,
            FormatScript::Subscript => xlsx::FormatScript::Subscript,
        }
    }
}
