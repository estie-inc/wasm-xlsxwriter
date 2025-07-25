
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex, MutexGuard};
use wasm_bindgen::prelude::*;
use crate::impl_method;

/// The `Format` struct is used to define cell formatting for data in a
/// worksheet.
/// 
/// The properties of a cell that can be formatted include: fonts, colors,
/// patterns, borders, alignment and number formatting.
/// 
/// <img src="https://rustxlsxwriter.github.io/images/format_intro.png">
/// 
/// The output file above was created with the following code:
/// 
/// ```
/// # // This code is available in examples/doc_format_intro.rs
/// #
/// use rust_xlsxwriter::{Format, Workbook, FormatBorder, Color, XlsxError};
/// 
/// fn main() -> Result<(), XlsxError> {
/// // Create a new Excel file object.
/// let mut workbook = Workbook::new();
/// 
/// // Add a worksheet.
/// let worksheet = workbook.add_worksheet();
/// 
/// // Make the first column wider for clarity.
/// worksheet.set_column_width(0, 14)?;
/// 
/// // Create some sample formats to display
/// let format1 = Format::new().set_font_name("Arial");
/// worksheet.write_with_format(0, 0, "Fonts", &format1)?;
/// 
/// let format2 = Format::new().set_font_name("Algerian").set_font_size(14);
/// worksheet.write_with_format(1, 0, "Fonts", &format2)?;
/// 
/// let format3 = Format::new().set_font_name("Comic Sans MS");
/// worksheet.write_with_format(2, 0, "Fonts", &format3)?;
/// 
/// let format4 = Format::new().set_font_name("Edwardian Script ITC");
/// worksheet.write_with_format(3, 0, "Fonts", &format4)?;
/// 
/// let format5 = Format::new().set_font_color(Color::Red);
/// worksheet.write_with_format(4, 0, "Font color", &format5)?;
/// 
/// let format6 = Format::new().set_background_color(Color::RGB(0xDAA520));
/// worksheet.write_with_format(5, 0, "Fills", &format6)?;
/// 
/// let format7 = Format::new().set_border(FormatBorder::Thin);
/// worksheet.write_with_format(6, 0, "Borders", &format7)?;
/// 
/// let format8 = Format::new().set_bold();
/// worksheet.write_with_format(7, 0, "Bold", &format8)?;
/// 
/// let format9 = Format::new().set_italic();
/// worksheet.write_with_format(8, 0, "Italic", &format9)?;
/// 
/// let format10 = Format::new().set_bold().set_italic();
/// worksheet.write_with_format(9, 0, "Bold and Italic", &format10)?;
/// 
/// workbook.save("formats.xlsx")?;
/// 
/// Ok(())
/// }
/// ```
/// 
/// # Contents
/// 
/// - [Creating and using a Format object](#creating-and-using-a-format-object)
/// - [Format methods and Format
/// properties](#format-methods-and-format-properties)
/// - [Format Colors](#format-colors)
/// - [Format Defaults](#format-defaults)
/// - [Row and Column Formats](#row-and-column-formats)
/// - [Number Format Categories](#number-format-categories)
/// - [Number Formats in Different
/// Locales](#number-formats-in-different-locales)
/// - [API](#implementations)
/// 
/// # Creating and using a Format object
/// 
/// Formats are created by calling the `Format::new()` method and properties as
/// set using the various methods shown is this section of the document. Once
/// the Format has been created it can be passed to one of the worksheet
/// `write_*()` methods. Multiple properties can be set by chaining them
/// together:
/// 
/// ```
/// # // This code is available in examples/doc_format_create.rs
/// #
/// # use rust_xlsxwriter::{Format, Workbook, Color, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     // Create a new Excel file object.
/// #     let mut workbook = Workbook::new();
/// #
/// #     // Add a worksheet.
/// #     let worksheet = workbook.add_worksheet();
/// #
/// // Create a new format and set some properties.
/// let format = Format::new()
/// .set_bold()
/// .set_italic()
/// .set_font_color(Color::Red);
/// 
/// worksheet.write_with_format(0, 0, "Hello", &format)?;
/// 
/// #     workbook.save("formats.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
/// 
/// <img src="https://rustxlsxwriter.github.io/images/format_create.png">
/// 
/// Formats can be cloned in the usual way:
/// 
/// ```
/// # // This code is available in examples/doc_format_clone.rs
/// #
/// # use rust_xlsxwriter::{Format, Workbook, Color, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     // Create a new Excel file object.
/// #     let mut workbook = Workbook::new();
/// #
/// #     // Add a worksheet.
/// #     let worksheet = workbook.add_worksheet();
/// #
/// #     // Create a new format and set some properties.
/// let format1 = Format::new()
/// .set_bold();
/// 
/// // Clone a new format and set some properties.
/// let format2 = format1.clone()
/// .set_font_color(Color::Blue);
/// 
/// worksheet.write_with_format(0, 0, "Hello", &format1)?;
/// worksheet.write_with_format(1, 0, "Hello", &format2)?;
/// 
/// #     workbook.save("formats.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
/// 
/// <img src="https://rustxlsxwriter.github.io/images/format_clone.png">
/// 
/// You can also merge two formats into a new combined format using the
/// [`Format::merge()`] method, as shown in the example below. The example also
/// demonstrates that properties in the primary format take precedence.
/// 
/// ```
/// # // This code is available in examples/doc_format_merge2.rs
/// #
/// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     // Create a new Excel file object.
/// #     let mut workbook = Workbook::new();
/// #
/// #     // Add a worksheet.
/// #     let worksheet = workbook.add_worksheet();
/// #
/// // Add some formats.
/// let format1 = Format::new().set_font_color("006100").set_bold();
/// let format2 = Format::new().set_font_color("9C0006").set_italic();
/// 
/// // Create new formats based on a merge of two formats.
/// let merged1 = format1.merge(&format2);
/// let merged2 = format2.merge(&format1);
/// 
/// // Write some strings with the formats.
/// worksheet.write_with_format(0, 0, "Format 1: green and bold", &format1)?;
/// worksheet.write_with_format(1, 0, "Format 2: red and italic", &format2)?;
/// worksheet.write_with_format(3, 0, "Merged 2 into 1", &merged1)?;
/// worksheet.write_with_format(4, 0, "Merged 1 into 2", &merged2)?;
/// #
/// #     // Autofit for clarity.
/// #     worksheet.autofit();
/// #
/// #     // Save the file.
/// #     workbook.save("formats.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
/// 
/// Output file:
/// 
/// <img src="https://rustxlsxwriter.github.io/images/format_merge2.png">
/// 
/// 
/// 
/// 
/// # Format methods and Format properties
/// 
/// The following table shows the Excel format categories, in the order shown in
/// the Excel "Format Cell" dialog, and the equivalent `rust_xlsxwriter` Format
/// method:
/// 
/// | Category        | Description           |  Method Name                             |
/// | :-------------- | :-------------------- |  :-------------------------------------- |
/// | **Number**      | Numeric format        |  [`Format::setNumFormat()`]            |
/// | **Alignment**   | Horizontal align      |  [`Format::setAlign()`]                 |
/// |                 | Vertical align        |  [`Format::setAlign()`]                 |
/// |                 | Rotation              |  [`Format::setRotation()`]              |
/// |                 | Text wrap             |  [`Format::setTextWrap()`]             |
/// |                 | Indentation           |  [`Format::setIndent()`]                |
/// |                 | Reading direction     |  [`Format::setReadingDirection()`]     |
/// |                 | Shrink to fit         |  [`Format::setShrink()`]                |
/// | **Font**        | Font type             |  [`Format::setFontName()`]             |
/// |                 | Font size             |  [`Format::setFontSize()`]             |
/// |                 | Font color            |  [`Format::setFontColor()`]            |
/// |                 | Bold                  |  [`Format::setBold()`]                  |
/// |                 | Italic                |  [`Format::setItalic()`]                |
/// |                 | Underline             |  [`Format::setUnderline()`]             |
/// |                 | Strikethrough         |  [`Format::setFontStrikethrough()`]    |
/// |                 | Super/Subscript       |  [`Format::setFontScript()`]           |
/// | **Border**      | Cell border           |  [`Format::setBorder()`]                |
/// |                 | Bottom border         |  [`Format::setBorderBottom()`]         |
/// |                 | Top border            |  [`Format::setBorderTop()`]            |
/// |                 | Left border           |  [`Format::setBorderLeft()`]           |
/// |                 | Right border          |  [`Format::setBorderRight()`]          |
/// |                 | Border color          |  [`Format::setBorderColor()`]          |
/// |                 | Bottom color          |  [`Format::setBorderBottomColor()`]   |
/// |                 | Top color             |  [`Format::setBorderTopColor()`]      |
/// |                 | Left color            |  [`Format::setBorderLeftColor()`]     |
/// |                 | Right color           |  [`Format::setBorderRightColor()`]    |
/// |                 | Diagonal border       |  [`Format::setBorderDiagonal()`]       |
/// |                 | Diagonal border color |  [`Format::setBorderDiagonalColor()`] |
/// |                 | Diagonal border type  |  [`Format::setBorderDiagonalType()`]  |
/// | **Fill**        | Cell pattern          |  [`Format::setPattern()`]               |
/// |                 | Background color      |  [`Format::setBackgroundColor()`]      |
/// |                 | Foreground color      |  [`Format::setForegroundColor()`]      |
/// | **Protection**  | Unlock cells          |  [`Format::setUnlocked()`]              |
/// |                 | Hide formulas         |  [`Format::setHidden()`]                |
/// 
/// # Format Colors
/// 
/// Format property colors are specified by using the [`Color`] enum with a Html
/// style RGB integer value or a limited number of defined colors:
/// 
/// ```
/// # // This code is available in examples/doc_enum_Color.rs
/// #
/// # use rust_xlsxwriter::{Format, Workbook, Color, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// // Create a new Excel file object.
/// let mut workbook = Workbook::new();
/// 
/// let format1 = Format::new().set_font_color(Color::Red);
/// let format2 = Format::new().set_font_color(Color::Green);
/// let format3 = Format::new().set_font_color(Color::RGB(0x4F026A));
/// let format4 = Format::new().set_font_color(Color::RGB(0x73CC5F));
/// let format5 = Format::new().set_font_color(Color::RGB(0xFFACFF));
/// let format6 = Format::new().set_font_color(Color::RGB(0xCC7E16));
/// 
/// let worksheet = workbook.add_worksheet();
/// worksheet.write_with_format(0, 0, "Red", &format1)?;
/// worksheet.write_with_format(1, 0, "Green", &format2)?;
/// worksheet.write_with_format(2, 0, "#4F026A", &format3)?;
/// worksheet.write_with_format(3, 0, "#73CC5F", &format4)?;
/// worksheet.write_with_format(4, 0, "#FFACFF", &format5)?;
/// worksheet.write_with_format(5, 0, "#CC7E16", &format6)?;
/// 
/// #     workbook.save("colors.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
/// 
/// <img src="https://rustxlsxwriter.github.io/images/enum_xlsxcolor.png">
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
/// ```
/// # // This code is available in examples/doc_format_default.rs
/// #
/// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     // Create a new Excel file object.
/// #     let mut workbook = Workbook::new();
/// #
/// #     // Add a worksheet.
/// #     let worksheet = workbook.add_worksheet();
/// #
/// // Create a new default format.
/// let format = Format::default();
/// 
/// // These methods calls are equivalent.
/// worksheet.write(0, 0, "Hello")?;
/// worksheet.write_with_format(1, 0, "Hello", &format)?;
/// #
/// #     workbook.save("formats.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
/// 
/// <img src="https://rustxlsxwriter.github.io/images/format_default.png">
/// 
/// 
/// # Row and Column Formats
/// 
/// Formatting can be applied to rows and columns in a worksheet using the
/// [`Worksheet::setRowFormat()`](crate::Worksheet::set_row_format) and
/// [`Worksheet::setColumnFormat()`](crate::Worksheet::set_column_format)
/// methods. The `rust_xlsxwriter` library ensures that the row and column
/// formats are applied to any cells in the row or column that contain data but
/// don't have an explicit format. It also ensures that the cell at the
/// intersection of a row and column format inherits the properties of both
/// formats, like in Excel:
/// 
/// ```
/// # // This code is available in examples/doc_format_merge3.rs
/// #
/// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     // Create a new Excel file object.
/// #     let mut workbook = Workbook::new();
/// #
/// #     // Add a worksheet.
/// #     let worksheet = workbook.add_worksheet();
/// #
/// // Add some formats.
/// let red = Format::new().set_font_color("9C0006");
/// let bold = Format::new().set_bold();
/// 
/// // Set some row and column formats.
/// worksheet.set_row_format(2, &bold)?;
/// worksheet.set_column_format(2, &red)?;
/// 
/// // Write some strings without explicit formats.
/// worksheet.write(0, 2, "C1")?; // Red.
/// worksheet.write(2, 0, "A3")?; // Bold.
/// worksheet.write(2, 2, "C3")?; // Bold and red.
/// #
/// #     // Save the file.
/// #     workbook.save("formats.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
/// 
/// <img src="https://rustxlsxwriter.github.io/images/format_merge3.png">
/// 
/// It should be noted that the intersection format (in cell C3 in the example
/// above) requires that the row and column formats are set before data is
/// written to the intersection cell. This restriction doesn't apply to
/// non-intersection row and column cells.
/// 
/// 
/// # Number Format Categories
/// 
/// The [`Format::setNumFormat()`] method is used to set the number format for
/// numbers used with
/// [`Worksheet::writeWithFormat()`](crate::Worksheet::write_with_format())
/// and
/// [`Worksheet::writeNumberWithFormat()`](crate::Worksheet::write_number_with_format()):
/// 
/// ```
/// # // This code is available in examples/doc_format_currency1.rs
/// 
/// use rust_xlsxwriter::{Format, Workbook, XlsxError};
/// 
/// fn main() -> Result<(), XlsxError> {
/// // Create a new Excel file object.
/// let mut workbook = Workbook::new();
/// 
/// // Add a worksheet.
/// let worksheet = workbook.add_worksheet();
/// 
/// // Add a format.
/// let currency_format = Format::new().set_num_format("$#,##0.00");
/// 
/// worksheet.write_with_format(0, 0, 1234.56, &currency_format)?;
/// 
/// workbook.save("currency_format.xlsx")?;
/// 
/// Ok(())
/// }
/// ```
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
/// ```
/// # // This code is available in examples/doc_format_currency2.rs
/// #
/// # use rust_xlsxwriter::{Format, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #    // Create a new Excel file object.
/// #    let mut workbook = Workbook::new();
/// #
/// #    // Add a worksheet.
/// #    let worksheet = workbook.add_worksheet();
/// #
/// #    // Add a format.
/// let currency_format = Format::new().set_num_format("[$$-409]#,##0.00");
/// 
/// worksheet.write_with_format(0, 0, 1234.56, &currency_format)?;
/// 
/// #   workbook.save("currency_format.xlsx")?;
/// #
/// #   Ok(())
/// # }
/// ```
/// 
/// That give us the following updated output. Note that the number category is
/// now shown as Currency:
/// 
/// <img src="https://rustxlsxwriter.github.io/images/format_currency4.png">
/// 
/// The same process can be used to find format strings for "Date" or
/// "Accountancy" formats.
/// 
/// # Number Formats in Different Locales
/// 
/// As shown in the previous section the `format.set_num_format()` method is
/// used to set the number format for `rust_xlsxwriter` formats. A common use
/// case is to set a number format with a "grouping/thousands" separator and a
/// "decimal" point:
/// 
/// ```
/// # // This code is available in examples/doc_format_locale.rs
/// #
/// use rust_xlsxwriter::{Format, Workbook, XlsxError};
/// 
/// fn main() -> Result<(), XlsxError> {
/// // Create a new Excel file object.
/// let mut workbook = Workbook::new();
/// 
/// 
/// // Add a worksheet.
/// let worksheet = workbook.add_worksheet();
/// 
/// // Add a format.
/// let currency_format = Format::new().set_num_format("#,##0.00");
/// 
/// worksheet.write_with_format(0, 0, 1234.56, &currency_format)?;
/// 
/// workbook.save("number_format.xlsx")?;
/// 
/// Ok(())
/// }
/// ```
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
#[derive(Debug, Clone)]
#[wasm_bindgen]
struct Format {
    pub(crate) inner: Arc::<Mutex::<xlsx::Format>>
}

#[wasm_bindgen]
impl Format {
    /// Create a new Format object.
    /// 
    /// Create a new Format object to use with worksheet formatting.
    #[wasm_bindgen(constructor)]
    pub fn new() -> Format {
        Format {inner: Arc::new(Mutex::new(xlsx::Format::new()))
        }
    }

    pub(crate) fn lock(&self) -> MutexGuard::<xlsx::Format> {
        (self.inner).lock().unwrap()
    }

    /// Deep clones a Format object.
    #[wasm_bindgen(js_name="clone")]
    pub fn deep_clone(&self) -> Format {
        let inner = (self.inner).lock().unwrap();
        Format {inner: Arc::new(Mutex::new(inner.clone()))
        }
    }

    /// Merge two formats into a new combined Format.
    /// 
    /// The method returns a new Format object that is a combination of two
    /// formats.
    /// 
    /// Precedence is given to the properties of the primary format that have
    /// been set. The primary format is the calling format. The order of
    /// precedence can be reversed by reversing the order of the primary and
    /// secondary formats, see the second example below.
    /// 
    /// # Parameters
    /// 
    /// - `other`: A Format object to merge with the primary Format.
    #[wasm_bindgen(js_name = "merge", skip_jsdoc)]
    pub fn merge(&self, other: Format) -> Format {
        impl_method!(self.merge(other));
    }

    /// Set the Format alignment properties.
    /// 
    /// This method is used to set the horizontal and vertical data alignment
    /// within a cell.
    /// 
    /// # Parameters
    /// 
    /// - `align`: The vertical and or horizontal alignment direction as
    /// defined by the [`FormatAlign`] enum.
    #[wasm_bindgen(js_name = "setAlign", skip_jsdoc)]
    pub fn set_align(&self, align: FormatAlign) -> Format {
        impl_method!(self.set_align(align.into()));
    }

    /// Set the Format pattern background color property.
    /// 
    /// The `set_background_color` method can be used to set the background
    /// color of a pattern. Patterns are defined via the [`Format::set_pattern`]
    /// method. If a pattern hasn't been defined then a solid fill pattern is
    /// used as the default.
    /// 
    /// # Parameters
    /// 
    /// - `color`: The background color property defined by a [`Color`] enum
    /// value or a type that can convert [`Into`] a [`Color`].
    #[wasm_bindgen(js_name = "setBackgroundColor", skip_jsdoc)]
    pub fn set_background_color(&self, color: Color) -> Format {
        impl_method!(self.set_background_color(color));
    }

    /// Set the bold property for a Format font.
    #[wasm_bindgen(js_name = "setBold", skip_jsdoc)]
    pub fn set_bold(&self) -> Format {
        impl_method!(self.set_bold());
    }

    /// Set the Format border property.
    /// 
    /// Set the cell border style. Individual border elements can be configured
    /// using the following methods with the same parameters:
    /// 
    /// - [`Format::setBorderTop()`]
    /// - [`Format::setBorderLeft()`]
    /// - [`Format::setBorderRight()`]
    /// - [`Format::setBorderColor()`]
    /// 
    /// # Parameters
    /// 
    /// - `border`: The border property as defined by a [`FormatBorder`] enum
    /// value.
    #[wasm_bindgen(js_name = "setBorder", skip_jsdoc)]
    pub fn set_border(&self, border: FormatBorder) -> Format {
        impl_method!(self.set_border(border.into()));
    }

    /// Set the cell bottom border style.
    /// 
    /// See [`Format::setBorder()`] for details.
    /// 
    /// # Parameters
    /// 
    /// - `border`: The border property as defined by a [`FormatBorder`] enum
    /// value.
    #[wasm_bindgen(js_name = "setBorderBottom", skip_jsdoc)]
    pub fn set_border_bottom(&self, border: FormatBorder) -> Format {
        impl_method!(self.set_border_bottom(border.into()));
    }

    /// Set the cell bottom border color.
    /// 
    /// See [`Format::setBorderColor()`] for details.
    /// 
    /// # Parameters
    /// 
    /// - `color`: The border color as defined by a [`Color`] enum value or a
    /// type that can convert [`Into`] a [`Color`].
    #[wasm_bindgen(js_name = "setBorderBottomColor", skip_jsdoc)]
    pub fn set_border_bottom_color(&self, color: Color) -> Format {
        impl_method!(self.set_border_bottom_color(color));
    }

    /// Set the Format border color property.
    /// 
    /// Set the cell border color. Individual border elements can be configured
    /// using the following methods with the same parameters:
    /// 
    /// - [`Format::setBorderTopColor()`]
    /// - [`Format::setBorderLeftColor()`]
    /// - [`Format::setBorderRightColor()`]
    /// - [`Format::setBorderBottomColor()`]
    /// 
    /// # Parameters
    /// 
    /// - `color`: The border color as defined by a [`Color`] enum value or
    /// a type that can convert [`Into`] a [`Color`].
    #[wasm_bindgen(js_name = "setBorderColor", skip_jsdoc)]
    pub fn set_border_color(&self, color: Color) -> Format {
        impl_method!(self.set_border_color(color));
    }

    /// Set the Format border diagonal property.
    /// 
    /// Set the cell border diagonal line style. This method should be used in
    /// conjunction with the [`Format::setBorderDiagonalType()`] method to
    /// set the diagonal type.
    /// 
    /// # Parameters
    /// 
    /// - `border`: The border property as defined by a [`FormatBorder`] enum
    /// value.
    #[wasm_bindgen(js_name = "setBorderDiagonal", skip_jsdoc)]
    pub fn set_border_diagonal(&self, border: FormatBorder) -> Format {
        impl_method!(self.set_border_diagonal(border.into()));
    }

    /// Set the cell diagonal border color.
    /// 
    /// See [`Format::setBorderDiagonal()`] for details.
    /// 
    /// # Parameters
    /// 
    /// - `color`: The border color as defined by a [`Color`] enum value or a
    /// type that can convert [`Into`] a [`Color`].
    #[wasm_bindgen(js_name = "setBorderDiagonalColor", skip_jsdoc)]
    pub fn set_border_diagonal_color(&self, color: Color) -> Format {
        impl_method!(self.set_border_diagonal_color(color));
    }

    /// Set the cell diagonal border direction type.
    /// 
    /// See [`Format::setBorderDiagonal()`] for details.
    /// 
    /// # Parameters
    /// 
    /// - `border_type`: The diagonal border type as defined by a
    /// [`FormatDiagonalBorder`] enum value.
    #[wasm_bindgen(js_name = "setBorderDiagonalType", skip_jsdoc)]
    pub fn set_border_diagonal_type(&self, border_type: FormatDiagonalBorder) -> Format {
        impl_method!(self.set_border_diagonal_type(border_type.into()));
    }

    /// Set the cell left border style.
    /// 
    /// See [`Format::setBorder()`] for details.
    /// 
    /// # Parameters
    /// 
    /// - `border`: The border property as defined by a [`FormatBorder`] enum
    /// value.
    #[wasm_bindgen(js_name = "setBorderLeft", skip_jsdoc)]
    pub fn set_border_left(&self, border: FormatBorder) -> Format {
        impl_method!(self.set_border_left(border.into()));
    }

    /// Set the cell left border color.
    /// 
    /// See [`Format::setBorderColor()`] for details.
    /// 
    /// # Parameters
    /// 
    /// - `color`: The border color as defined by a [`Color`] enum value or a
    /// type that can convert [`Into`] a [`Color`].
    #[wasm_bindgen(js_name = "setBorderLeftColor", skip_jsdoc)]
    pub fn set_border_left_color(&self, color: Color) -> Format {
        impl_method!(self.set_border_left_color(color));
    }

    /// Set the cell right border style.
    /// 
    /// See [`Format::setBorder()`] for details.
    /// 
    /// # Parameters
    /// 
    /// - `border`: The border property as defined by a [`FormatBorder`] enum
    /// value.
    #[wasm_bindgen(js_name = "setBorderRight", skip_jsdoc)]
    pub fn set_border_right(&self, border: FormatBorder) -> Format {
        impl_method!(self.set_border_right(border.into()));
    }

    /// Set the cell right border color.
    /// 
    /// See [`Format::setBorderColor()`] for details.
    /// 
    /// # Parameters
    /// 
    /// - `color`: The border color as defined by a [`Color`] enum value or a
    /// type that can convert [`Into`] a [`Color`].
    #[wasm_bindgen(js_name = "setBorderRightColor", skip_jsdoc)]
    pub fn set_border_right_color(&self, color: Color) -> Format {
        impl_method!(self.set_border_right_color(color));
    }

    /// Set the cell top border style.
    /// 
    /// See [`Format::setBorder()`] for details.
    /// 
    /// # Parameters
    /// 
    /// - `border`: The border property as defined by a [`FormatBorder`] enum
    /// value.
    #[wasm_bindgen(js_name = "setBorderTop", skip_jsdoc)]
    pub fn set_border_top(&self, border: FormatBorder) -> Format {
        impl_method!(self.set_border_top(border.into()));
    }

    /// Set the cell top border color.
    /// 
    /// See [`Format::setBorderColor()`] for details.
    /// 
    /// # Parameters
    /// 
    /// - `color`: The border color as defined by a [`Color`] enum value or a
    /// type that can convert [`Into`] a [`Color`].
    #[wasm_bindgen(js_name = "setBorderTopColor", skip_jsdoc)]
    pub fn set_border_top_color(&self, color: Color) -> Format {
        impl_method!(self.set_border_top_color(color));
    }

    /// Set the Format property to show a checkbox in a cell.
    /// 
    /// This format property can be used with a cell that contains a boolean
    /// value to display it as a checkbox. This property isn't required very
    /// often and it is generally easier to create a checkbox using the
    /// [`Worksheet::insertCheckbox()`](crate::Worksheet::insert_checkbox)
    /// method.
    #[wasm_bindgen(js_name = "setCheckbox", skip_jsdoc)]
    pub fn set_checkbox(&self) -> Format {
        impl_method!(self.set_checkbox());
    }

    /// Set the Format font character set property.
    /// 
    /// Set the font character. This function is implemented for completeness
    /// but is rarely used in practice.
    /// 
    /// # Parameters
    /// 
    /// - `font_charset`: The font character set property.
    #[wasm_bindgen(js_name = "setFontCharset", skip_jsdoc)]
    pub fn set_font_charset(&self, font_charset: u8) -> Format {
        impl_method!(self.set_font_charset(font_charset));
    }

    /// Set the color property for the Format font.
    /// 
    /// The `set_font_color()` method is used to set the font color in a cell.
    /// To set the color of a cell background use the `set_bg_color()` and
    /// `set_pattern()` methods.
    /// 
    /// # Parameters
    /// 
    /// - `color`: The font color property defined by a [`Color`] enum
    /// value.
    #[wasm_bindgen(js_name = "setFontColor", skip_jsdoc)]
    pub fn set_font_color(&self, color: Color) -> Format {
        impl_method!(self.set_font_color(color));
    }

    /// Set the Format font family property.
    /// 
    /// Set the font family. This is usually an integer in the range 1-4. This
    /// function is implemented for completeness but is rarely used in practice.
    /// 
    /// # Parameters
    /// 
    /// - `font_family`: The font family property.
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
    /// # Parameters
    /// 
    /// - `font_name`: The font name property.
    #[wasm_bindgen(js_name = "setFontName", skip_jsdoc)]
    pub fn set_font_name(&self, font_name: String) -> Format {
        impl_method!(self.set_font_name(font_name));
    }

    /// Set the Format font scheme property.
    /// 
    /// This function is implemented for completeness but is rarely used in
    /// practice.
    /// 
    /// # Parameters
    /// 
    /// - `font_scheme`: The font scheme property.
    #[wasm_bindgen(js_name = "setFontScheme", skip_jsdoc)]
    pub fn set_font_scheme(&self, font_scheme: String) -> Format {
        impl_method!(self.set_font_scheme(font_scheme));
    }

    /// Set the Format font super/subscript property.
    /// 
    /// This feature is generally only useful when using a font in a "rich"
    /// string. See
    /// [`write_rich_string()`](crate::Worksheet::write_rich_string).
    /// 
    /// # Parameters
    /// 
    /// - `font_script`: The font superscript or subscript property via a
    /// [`FormatScript`] enum.
    #[wasm_bindgen(js_name = "setFontScript", skip_jsdoc)]
    pub fn set_font_script(&self, font_script: FormatScript) -> Format {
        impl_method!(self.set_font_script(font_script.into()));
    }

    /// Set the Format font size property.
    /// 
    /// Set the font size of the cell format. The size is generally an integer
    /// value but Excel allows x.5 values (hence the property is a f64 or
    /// types that can convert [`Into`] a f64).
    /// 
    /// Excel adjusts the height of a row to accommodate the largest font size
    /// in the row.
    /// 
    /// # Parameters
    /// 
    /// - `font_size`: The font size property.
    #[wasm_bindgen(js_name = "setFontSize", skip_jsdoc)]
    pub fn set_font_size(&self, font_size: UnknownGeneric) -> Format {
        impl_method!(self.set_font_size(font_size));
    }

    /// Set the Format font strikethrough property.
    #[wasm_bindgen(js_name = "setFontStrikethrough", skip_jsdoc)]
    pub fn set_font_strikethrough(&self) -> Format {
        impl_method!(self.set_font_strikethrough());
    }

    /// Set the Format pattern foreground color property.
    /// 
    /// The `set_foreground_color` method can be used to set the
    /// foreground/pattern color of a pattern. Patterns are defined via the
    /// [`Format::set_pattern`] method.
    /// 
    /// # Parameters
    /// 
    /// - `color`: The foreground color property defined by a [`Color`] enum
    /// value or a type that can convert [`Into`] a [`Color`].
    #[wasm_bindgen(js_name = "setForegroundColor", skip_jsdoc)]
    pub fn set_foreground_color(&self, color: Color) -> Format {
        impl_method!(self.set_foreground_color(color));
    }

    /// Set the Format property to hide formulas in a cell.
    /// 
    /// This method can be used to hide a formula while still displaying its
    /// result. This is generally used to hide complex calculations from end
    /// users who are only interested in the result. It only has an effect if
    /// the worksheet has been protected using the
    /// [`Worksheet::protect()`](crate::Worksheet::protect) method.
    /// 
    /// See the example above.
    #[wasm_bindgen(js_name = "setHidden", skip_jsdoc)]
    pub fn set_hidden(&self) -> Format {
        impl_method!(self.set_hidden());
    }

    /// Set the hyperlink style.
    /// 
    /// Set the hyperlink style for use with urls. This is usually set
    /// automatically when writing urls without a format applied.
    #[wasm_bindgen(js_name = "setHyperlink", skip_jsdoc)]
    pub fn set_hyperlink(&self) -> Format {
        impl_method!(self.set_hyperlink());
    }

    /// Set the Format indent property.
    /// 
    /// This method can be used to indent text in a cell.
    /// 
    /// Indentation is a horizontal alignment property. It can be used in Excel
    /// in conjunction with the [`FormatAlign::Left`], [`FormatAlign::Right`]
    /// and [`FormatAlign::Distributed`] alignments. It will override any other
    /// horizontal properties that don't support indentation.
    /// 
    /// # Parameters
    /// 
    /// - `indent`: The indentation level for the cell.
    #[wasm_bindgen(js_name = "setIndent", skip_jsdoc)]
    pub fn set_indent(&self, indent: u8) -> Format {
        impl_method!(self.set_indent(indent));
    }

    /// Set the italic property for the Format font.
    #[wasm_bindgen(js_name = "setItalic", skip_jsdoc)]
    pub fn set_italic(&self) -> Format {
        impl_method!(self.set_italic());
    }

    /// Set the locked Format property back to its default "on" state.
    /// 
    /// The opposite of [`Format::setUnlocked()`].
    #[wasm_bindgen(js_name = "setLocked", skip_jsdoc)]
    pub fn set_locked(&self) -> Format {
        impl_method!(self.set_locked());
    }

    /// Set the number format for a Format.
    /// 
    /// This method is used to define the numerical format of a number in Excel.
    /// It controls whether a number is displayed as an integer, a floating-point
    /// number, a date, a currency value, or some other user-defined
    /// format.
    /// 
    /// See also [Number Format Categories] and [Number Formats in Different
    /// Locales].
    /// 
    /// [Number Format Categories]: crate::Format#number-format-categories
    /// [Number Formats in Different Locales]:
    /// crate::Format#number-formats-in-different-locales
    /// 
    /// # Parameters
    /// 
    /// - `num_format`: The number format property.
    #[wasm_bindgen(js_name = "setNumFormat", skip_jsdoc)]
    pub fn set_num_format(&self, num_format: String) -> Format {
        impl_method!(self.set_num_format(num_format));
    }

    /// Set the number format for a Format using a legacy format index.
    /// 
    /// This method is similar to [`Format::setNumFormat()`] except that it
    /// uses an index to a limited number of Excel's built-in, and legacy,
    /// number formats.
    /// 
    /// Unless you need to specifically access one of Excel's built-in number
    /// formats the [`Format::setNumFormat()`] method is a better solution.
    /// This method is mainly included for backward compatibility and
    /// completeness.
    /// 
    /// The Excel built-in number formats as shown in the table below:
    /// 
    /// | Index | Format String                                        |
    /// | :---- | :--------------------------------------------------- |
    /// | 1     | `0`                                                  |
    /// | 2     | `0.00`                                               |
    /// | 3     | `#,##0`                                              |
    /// | 4     | `#,##0.00`                                           |
    /// | 5     | `($#,##0_);($#,##0)`                                 |
    /// | 6     | `($#,##0_);[Red]($#,##0)`                            |
    /// | 7     | `($#,##0.00_);($#,##0.00)`                           |
    /// | 8     | `($#,##0.00_);[Red]($#,##0.00)`                      |
    /// | 9     | `0%`                                                 |
    /// | 10    | `0.00%`                                              |
    /// | 11    | `0.00E+00`                                           |
    /// | 12    | `# ?/?`                                              |
    /// | 13    | `# ??/??`                                            |
    /// | 14    | `m/d/yy`                                             |
    /// | 15    | `d-mmm-yy`                                           |
    /// | 16    | `d-mmm`                                              |
    /// | 17    | `mmm-yy`                                             |
    /// | 18    | `h:mm AM/PM`                                         |
    /// | 19    | `h:mm:ss AM/PM`                                      |
    /// | 20    | `h:mm`                                               |
    /// | 21    | `h:mm:ss`                                            |
    /// | 22    | `m/d/yy h:mm`                                        |
    /// | ...   | ...                                                  |
    /// | 37    | `(#,##0_);(#,##0)`                                   |
    /// | 38    | `(#,##0_);[Red](#,##0)`                              |
    /// | 39    | `(#,##0.00_);(#,##0.00)`                             |
    /// | 40    | `(#,##0.00_);[Red](#,##0.00)`                        |
    /// | 41    | `_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)`            |
    /// | 42    | `_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)`         |
    /// | 43    | `_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)`    |
    /// | 44    | `_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)` |
    /// | 45    | `mm:ss`                                              |
    /// | 46    | `[h]:mm:ss`                                          |
    /// | 47    | `mm:ss.0`                                            |
    /// | 48    | `##0.0E+0`                                           |
    /// | 49    | `@`                                                  |
    /// 
    /// Notes:
    /// 
    /// - Numeric formats 23 to 36 are not documented by Microsoft and may
    /// differ in international versions. The listed date and currency
    /// formats may also vary depending on system settings.
    /// - The dollar sign in the above format appears as the defined local
    /// currency symbol.
    /// - These formats can also be set via
    /// [`Format::setNumFormat()`].
    /// 
    /// # Parameters
    /// 
    /// - `num_format_index`: The index to one of the inbuilt formats shown in
    /// the table above.
    #[wasm_bindgen(js_name = "setNumFormatIndex", skip_jsdoc)]
    pub fn set_num_format_index(&self, num_format_index: u8) -> Format {
        impl_method!(self.set_num_format_index(num_format_index));
    }

    /// Set the Format pattern property.
    /// 
    /// Set the pattern for a cell. The most commonly used pattern is
    /// [`FormatPattern::Solid`].
    /// 
    /// To set the pattern colors see [`Format::setBackgroundColor()`] and
    /// [`Format::setForegroundColor()`].
    /// 
    /// # Parameters
    /// 
    /// - `pattern`: The pattern property defined by a [`FormatPattern`] enum
    /// value.
    #[wasm_bindgen(js_name = "setPattern", skip_jsdoc)]
    pub fn set_pattern(&self, pattern: FormatPattern) -> Format {
        impl_method!(self.set_pattern(pattern.into()));
    }

    /// Set the `quote_prefix` property for a Format.
    /// 
    /// Set the quote prefix property of a format to ensure a string is treated
    /// as a string after editing. This is the same as prefixing the string with
    /// a single quote in Excel. You don't need to add the quote to the string
    /// but you do need to add the format.
    #[wasm_bindgen(js_name = "setQuotePrefix", skip_jsdoc)]
    pub fn set_quote_prefix(&self) -> Format {
        impl_method!(self.set_quote_prefix());
    }

    /// Set the Format text reading order property.
    /// 
    /// Set the text reading direction. This is useful when creating Arabic,
    /// Hebrew or other near or far eastern worksheets. It can be used in
    /// conjunction with the Worksheet
    /// [`set_right_to_left`](crate::Worksheet::set_right_to_left()) method
    /// which changes the cell display direction of the worksheet.
    /// 
    /// # Parameters
    /// 
    /// - `reading_direction`: The reading order property, should be 0, 1, or
    /// 2, where these values refer to:
    /// 
    /// 0. The reading direction is determined heuristically by Excel
    /// depending on the text. This is the default option.
    /// 1. The text is displayed Left-to-Right, like English.
    /// 2. The text is displayed Right-to-Left, like Hebrew or Arabic.
    #[wasm_bindgen(js_name = "setReadingDirection", skip_jsdoc)]
    pub fn set_reading_direction(&self, reading_direction: u8) -> Format {
        impl_method!(self.set_reading_direction(reading_direction));
    }

    /// Set the Format rotation property.
    /// 
    /// Set the rotation angle of the text in a cell. The rotation can be any
    /// angle in the range -90 to 90 degrees, or 270 to indicate text where the
    /// letters run from top to bottom.
    /// 
    /// # Parameters
    /// 
    /// - `rotation`: The rotation angle.
    #[wasm_bindgen(js_name = "setRotation", skip_jsdoc)]
    pub fn set_rotation(&self, rotation: i16) -> Format {
        impl_method!(self.set_rotation(rotation));
    }

    /// Set the Format shrink property.
    /// 
    /// This method can be used to shrink text so that it fits in a cell
    #[wasm_bindgen(js_name = "setShrink", skip_jsdoc)]
    pub fn set_shrink(&self) -> Format {
        impl_method!(self.set_shrink());
    }

    /// Set the Format text wrap property.
    /// 
    /// This method is used to turn on automatic text wrapping for text in a
    /// cell. If you wish to control where the string is wrapped you can add
    /// newlines to the text (see the example below).
    /// 
    /// Excel generally adjusts the height of the cell to fit the wrapped text
    /// unless a explicit row height has be set via
    /// [`Worksheet::setRowHeight()`](crate::Worksheet::set_row_height()).
    #[wasm_bindgen(js_name = "setTextWrap", skip_jsdoc)]
    pub fn set_text_wrap(&self) -> Format {
        impl_method!(self.set_text_wrap());
    }

    /// Set the underline properties for a format.
    /// 
    /// The difference between a normal underline and an "accounting" underline
    /// is that a normal underline only underlines the text/number in a cell
    /// whereas an accounting underline underlines the entire cell width.
    /// 
    /// # Parameters
    /// 
    /// - `underline`: The underline type defined by a [`FormatUnderline`] enum
    /// value.
    #[wasm_bindgen(js_name = "setUnderline", skip_jsdoc)]
    pub fn set_underline(&self, underline: FormatUnderline) -> Format {
        impl_method!(self.set_underline(underline.into()));
    }

    /// Set the Format cell unlocked state.
    /// 
    /// This method can be used to allow modification of a cell in a protected
    /// worksheet. In Excel, cell locking is turned on by default for all cells.
    /// However, it only has an effect if the worksheet has been protected using
    /// the [`Worksheet::protect()`](crate::Worksheet::protect) method.
    #[wasm_bindgen(js_name = "setUnlocked", skip_jsdoc)]
    pub fn set_unlocked(&self) -> Format {
        impl_method!(self.set_unlocked());
    }

    /// Unset the bold Format property back to its default "off" state.
    /// 
    /// The opposite of [`Format::setBold()`].
    #[wasm_bindgen(js_name = "unsetBold", skip_jsdoc)]
    pub fn unset_bold(&self) -> Format {
        impl_method!(self.unset_bold());
    }

    /// Unset the `checkbox` Format property back to its default "off" state.
    /// 
    /// The opposite of [`Format::setCheckbox()`].
    #[wasm_bindgen(js_name = "unsetCheckbox", skip_jsdoc)]
    pub fn unset_checkbox(&self) -> Format {
        impl_method!(self.unset_checkbox());
    }

    /// Unset the font strikethrough Format property back to its default "off" state.
    /// 
    /// The opposite of [`Format::setFontStrikethrough()`].
    #[wasm_bindgen(js_name = "unsetFontStrikethrough", skip_jsdoc)]
    pub fn unset_font_strikethrough(&self) -> Format {
        impl_method!(self.unset_font_strikethrough());
    }

    /// Unset the hidden Format property back to its default "off" state.
    /// 
    /// The opposite of [`Format::setHidden()`].
    #[wasm_bindgen(js_name = "unsetHidden", skip_jsdoc)]
    pub fn unset_hidden(&self) -> Format {
        impl_method!(self.unset_hidden());
    }

    /// Unset the hyperlink style. Doesn't reset the other properties.
    #[wasm_bindgen(js_name = "unsetHyperlinkStyle", skip_jsdoc)]
    pub fn unset_hyperlink_style(&self) -> Format {
        impl_method!(self.unset_hyperlink_style());
    }

    /// Unset the italic Format property back to its default "off" state.
    /// 
    /// The opposite of [`Format::setItalic()`].
    #[wasm_bindgen(js_name = "unsetItalic", skip_jsdoc)]
    pub fn unset_italic(&self) -> Format {
        impl_method!(self.unset_italic());
    }

    /// Unset the `quote_prefix` Format property back to its default "off" state.
    /// 
    /// The opposite of [`Format::setQuotePrefix()`].
    #[wasm_bindgen(js_name = "unsetQuotePrefix", skip_jsdoc)]
    pub fn unset_quote_prefix(&self) -> Format {
        impl_method!(self.unset_quote_prefix());
    }

    /// Unset the shrink Format property back to its default "off" state.
    /// 
    /// The opposite of [`Format::setShrink()`].
    #[wasm_bindgen(js_name = "unsetShrink", skip_jsdoc)]
    pub fn unset_shrink(&self) -> Format {
        impl_method!(self.unset_shrink());
    }

    /// Unset the text wrap Format property back to its default "off" state.
    /// 
    /// The opposite of [`Format::setTextWrap()`].
    #[wasm_bindgen(js_name = "unsetTextWrap", skip_jsdoc)]
    pub fn unset_text_wrap(&self) -> Format {
        impl_method!(self.unset_text_wrap());
    }
}

