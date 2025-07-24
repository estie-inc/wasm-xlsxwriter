
use std::sync::{Arc, Mutex, MutexGuard};
use wasm_bindgen::prelude::*;
use rust_xlsxwriter::xlsx as xlsx;

#[Derive(Debug , Clone)]
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
        Format {inner: Arc::new(Mutex::new(inner))
        }
    }

    pub(crate) fn lock(&self) -> MutexGuard::<xlsx::Format> {
        (self.inner).lock().unwrap()
    }

    /// Deep clones a Format object.
    #[wasm_bindgen(js_name = "clone")]
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
    #[wasm_bindgen(js_name = "merge" , skip_jsdoc)]
    pub fn merge(&self, other: Format) -> Format {
        impl_method!(self . merge ( other ))
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
    #[wasm_bindgen(js_name = "setAlign" , skip_jsdoc)]
    pub fn set_align(&self, align: FormatAlign) -> Format {
        impl_method!(self . set_align ( align ))
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
    #[wasm_bindgen(js_name = "setBackgroundColor" , skip_jsdoc)]
    pub fn set_background_color(&self, color: Unknown) -> Format {
        impl_method!(self . set_background_color ( color ))
    }

    /// Set the bold property for a Format font.
    #[wasm_bindgen(js_name = "setBold" , skip_jsdoc)]
    pub fn set_bold(&self) -> Format {
        impl_method!(self . set_bold ( ))
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
    #[wasm_bindgen(js_name = "setBorder" , skip_jsdoc)]
    pub fn set_border(&self, border: FormatBorder) -> Format {
        impl_method!(self . set_border ( border ))
    }

    /// Set the cell bottom border style.
    /// 
    /// See [`Format::setBorder()`] for details.
    /// 
    /// # Parameters
    /// 
    /// - `border`: The border property as defined by a [`FormatBorder`] enum
    /// value.
    #[wasm_bindgen(js_name = "setBorderBottom" , skip_jsdoc)]
    pub fn set_border_bottom(&self, border: FormatBorder) -> Format {
        impl_method!(self . set_border_bottom ( border ))
    }

    /// Set the cell bottom border color.
    /// 
    /// See [`Format::setBorderColor()`] for details.
    /// 
    /// # Parameters
    /// 
    /// - `color`: The border color as defined by a [`Color`] enum value or a
    /// type that can convert [`Into`] a [`Color`].
    #[wasm_bindgen(js_name = "setBorderBottomColor" , skip_jsdoc)]
    pub fn set_border_bottom_color(&self, color: Unknown) -> Format {
        impl_method!(self . set_border_bottom_color ( color ))
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
    #[wasm_bindgen(js_name = "setBorderColor" , skip_jsdoc)]
    pub fn set_border_color(&self, color: Unknown) -> Format {
        impl_method!(self . set_border_color ( color ))
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
    #[wasm_bindgen(js_name = "setBorderDiagonal" , skip_jsdoc)]
    pub fn set_border_diagonal(&self, border: FormatBorder) -> Format {
        impl_method!(self . set_border_diagonal ( border ))
    }

    /// Set the cell diagonal border color.
    /// 
    /// See [`Format::setBorderDiagonal()`] for details.
    /// 
    /// # Parameters
    /// 
    /// - `color`: The border color as defined by a [`Color`] enum value or a
    /// type that can convert [`Into`] a [`Color`].
    #[wasm_bindgen(js_name = "setBorderDiagonalColor" , skip_jsdoc)]
    pub fn set_border_diagonal_color(&self, color: Unknown) -> Format {
        impl_method!(self . set_border_diagonal_color ( color ))
    }

    /// Set the cell diagonal border direction type.
    /// 
    /// See [`Format::setBorderDiagonal()`] for details.
    /// 
    /// # Parameters
    /// 
    /// - `border_type`: The diagonal border type as defined by a
    /// [`FormatDiagonalBorder`] enum value.
    #[wasm_bindgen(js_name = "setBorderDiagonalType" , skip_jsdoc)]
    pub fn set_border_diagonal_type(&self, border_type: FormatDiagonalBorder) -> Format {
        impl_method!(self . set_border_diagonal_type ( border_type ))
    }

    /// Set the cell left border style.
    /// 
    /// See [`Format::setBorder()`] for details.
    /// 
    /// # Parameters
    /// 
    /// - `border`: The border property as defined by a [`FormatBorder`] enum
    /// value.
    #[wasm_bindgen(js_name = "setBorderLeft" , skip_jsdoc)]
    pub fn set_border_left(&self, border: FormatBorder) -> Format {
        impl_method!(self . set_border_left ( border ))
    }

    /// Set the cell left border color.
    /// 
    /// See [`Format::setBorderColor()`] for details.
    /// 
    /// # Parameters
    /// 
    /// - `color`: The border color as defined by a [`Color`] enum value or a
    /// type that can convert [`Into`] a [`Color`].
    #[wasm_bindgen(js_name = "setBorderLeftColor" , skip_jsdoc)]
    pub fn set_border_left_color(&self, color: Unknown) -> Format {
        impl_method!(self . set_border_left_color ( color ))
    }

    /// Set the cell right border style.
    /// 
    /// See [`Format::setBorder()`] for details.
    /// 
    /// # Parameters
    /// 
    /// - `border`: The border property as defined by a [`FormatBorder`] enum
    /// value.
    #[wasm_bindgen(js_name = "setBorderRight" , skip_jsdoc)]
    pub fn set_border_right(&self, border: FormatBorder) -> Format {
        impl_method!(self . set_border_right ( border ))
    }

    /// Set the cell right border color.
    /// 
    /// See [`Format::setBorderColor()`] for details.
    /// 
    /// # Parameters
    /// 
    /// - `color`: The border color as defined by a [`Color`] enum value or a
    /// type that can convert [`Into`] a [`Color`].
    #[wasm_bindgen(js_name = "setBorderRightColor" , skip_jsdoc)]
    pub fn set_border_right_color(&self, color: Unknown) -> Format {
        impl_method!(self . set_border_right_color ( color ))
    }

    /// Set the cell top border style.
    /// 
    /// See [`Format::setBorder()`] for details.
    /// 
    /// # Parameters
    /// 
    /// - `border`: The border property as defined by a [`FormatBorder`] enum
    /// value.
    #[wasm_bindgen(js_name = "setBorderTop" , skip_jsdoc)]
    pub fn set_border_top(&self, border: FormatBorder) -> Format {
        impl_method!(self . set_border_top ( border ))
    }

    /// Set the cell top border color.
    /// 
    /// See [`Format::setBorderColor()`] for details.
    /// 
    /// # Parameters
    /// 
    /// - `color`: The border color as defined by a [`Color`] enum value or a
    /// type that can convert [`Into`] a [`Color`].
    #[wasm_bindgen(js_name = "setBorderTopColor" , skip_jsdoc)]
    pub fn set_border_top_color(&self, color: Unknown) -> Format {
        impl_method!(self . set_border_top_color ( color ))
    }

    /// Set the Format property to show a checkbox in a cell.
    /// 
    /// This format property can be used with a cell that contains a boolean
    /// value to display it as a checkbox. This property isn't required very
    /// often and it is generally easier to create a checkbox using the
    /// [`Worksheet::insertCheckbox()`](crate::Worksheet::insert_checkbox)
    /// method.
    #[wasm_bindgen(js_name = "setCheckbox" , skip_jsdoc)]
    pub fn set_checkbox(&self) -> Format {
        impl_method!(self . set_checkbox ( ))
    }

    /// Set the Format font character set property.
    /// 
    /// Set the font character. This function is implemented for completeness
    /// but is rarely used in practice.
    /// 
    /// # Parameters
    /// 
    /// - `font_charset`: The font character set property.
    #[wasm_bindgen(js_name = "setFontCharset" , skip_jsdoc)]
    pub fn set_font_charset(&self, font_charset: u8) -> Format {
        impl_method!(self . set_font_charset ( font_charset ))
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
    #[wasm_bindgen(js_name = "setFontColor" , skip_jsdoc)]
    pub fn set_font_color(&self, color: Unknown) -> Format {
        impl_method!(self . set_font_color ( color ))
    }

    /// Set the Format font family property.
    /// 
    /// Set the font family. This is usually an integer in the range 1-4. This
    /// function is implemented for completeness but is rarely used in practice.
    /// 
    /// # Parameters
    /// 
    /// - `font_family`: The font family property.
    #[wasm_bindgen(js_name = "setFontFamily" , skip_jsdoc)]
    pub fn set_font_family(&self, font_family: u8) -> Format {
        impl_method!(self . set_font_family ( font_family ))
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
    #[wasm_bindgen(js_name = "setFontName" , skip_jsdoc)]
    pub fn set_font_name(&self, font_name: Unknown) -> Format {
        impl_method!(self . set_font_name ( font_name ))
    }

    /// Set the Format font scheme property.
    /// 
    /// This function is implemented for completeness but is rarely used in
    /// practice.
    /// 
    /// # Parameters
    /// 
    /// - `font_scheme`: The font scheme property.
    #[wasm_bindgen(js_name = "setFontScheme" , skip_jsdoc)]
    pub fn set_font_scheme(&self, font_scheme: Unknown) -> Format {
        impl_method!(self . set_font_scheme ( font_scheme ))
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
    #[wasm_bindgen(js_name = "setFontScript" , skip_jsdoc)]
    pub fn set_font_script(&self, font_script: FormatScript) -> Format {
        impl_method!(self . set_font_script ( font_script ))
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
    #[wasm_bindgen(js_name = "setFontSize" , skip_jsdoc)]
    pub fn set_font_size(&self, font_size: Unknown) -> Format {
        impl_method!(self . set_font_size ( font_size ))
    }

    /// Set the Format font strikethrough property.
    #[wasm_bindgen(js_name = "setFontStrikethrough" , skip_jsdoc)]
    pub fn set_font_strikethrough(&self) -> Format {
        impl_method!(self . set_font_strikethrough ( ))
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
    #[wasm_bindgen(js_name = "setForegroundColor" , skip_jsdoc)]
    pub fn set_foreground_color(&self, color: Unknown) -> Format {
        impl_method!(self . set_foreground_color ( color ))
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
    #[wasm_bindgen(js_name = "setHidden" , skip_jsdoc)]
    pub fn set_hidden(&self) -> Format {
        impl_method!(self . set_hidden ( ))
    }

    /// Set the hyperlink style.
    /// 
    /// Set the hyperlink style for use with urls. This is usually set
    /// automatically when writing urls without a format applied.
    #[wasm_bindgen(js_name = "setHyperlink" , skip_jsdoc)]
    pub fn set_hyperlink(&self) -> Format {
        impl_method!(self . set_hyperlink ( ))
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
    #[wasm_bindgen(js_name = "setIndent" , skip_jsdoc)]
    pub fn set_indent(&self, indent: u8) -> Format {
        impl_method!(self . set_indent ( indent ))
    }

    /// Set the italic property for the Format font.
    #[wasm_bindgen(js_name = "setItalic" , skip_jsdoc)]
    pub fn set_italic(&self) -> Format {
        impl_method!(self . set_italic ( ))
    }

    /// Set the locked Format property back to its default "on" state.
    /// 
    /// The opposite of [`Format::setUnlocked()`].
    #[wasm_bindgen(js_name = "setLocked" , skip_jsdoc)]
    pub fn set_locked(&self) -> Format {
        impl_method!(self . set_locked ( ))
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
    #[wasm_bindgen(js_name = "setNumFormat" , skip_jsdoc)]
    pub fn set_num_format(&self, num_format: Unknown) -> Format {
        impl_method!(self . set_num_format ( num_format ))
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
    #[wasm_bindgen(js_name = "setNumFormatIndex" , skip_jsdoc)]
    pub fn set_num_format_index(&self, num_format_index: u8) -> Format {
        impl_method!(self . set_num_format_index ( num_format_index ))
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
    #[wasm_bindgen(js_name = "setPattern" , skip_jsdoc)]
    pub fn set_pattern(&self, pattern: FormatPattern) -> Format {
        impl_method!(self . set_pattern ( pattern ))
    }

    /// Set the `quote_prefix` property for a Format.
    /// 
    /// Set the quote prefix property of a format to ensure a string is treated
    /// as a string after editing. This is the same as prefixing the string with
    /// a single quote in Excel. You don't need to add the quote to the string
    /// but you do need to add the format.
    #[wasm_bindgen(js_name = "setQuotePrefix" , skip_jsdoc)]
    pub fn set_quote_prefix(&self) -> Format {
        impl_method!(self . set_quote_prefix ( ))
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
    #[wasm_bindgen(js_name = "setReadingDirection" , skip_jsdoc)]
    pub fn set_reading_direction(&self, reading_direction: u8) -> Format {
        impl_method!(self . set_reading_direction ( reading_direction ))
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
    #[wasm_bindgen(js_name = "setRotation" , skip_jsdoc)]
    pub fn set_rotation(&self, rotation: i16) -> Format {
        impl_method!(self . set_rotation ( rotation ))
    }

    /// Set the Format shrink property.
    /// 
    /// This method can be used to shrink text so that it fits in a cell
    #[wasm_bindgen(js_name = "setShrink" , skip_jsdoc)]
    pub fn set_shrink(&self) -> Format {
        impl_method!(self . set_shrink ( ))
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
    #[wasm_bindgen(js_name = "setTextWrap" , skip_jsdoc)]
    pub fn set_text_wrap(&self) -> Format {
        impl_method!(self . set_text_wrap ( ))
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
    #[wasm_bindgen(js_name = "setUnderline" , skip_jsdoc)]
    pub fn set_underline(&self, underline: FormatUnderline) -> Format {
        impl_method!(self . set_underline ( underline ))
    }

    /// Set the Format cell unlocked state.
    /// 
    /// This method can be used to allow modification of a cell in a protected
    /// worksheet. In Excel, cell locking is turned on by default for all cells.
    /// However, it only has an effect if the worksheet has been protected using
    /// the [`Worksheet::protect()`](crate::Worksheet::protect) method.
    #[wasm_bindgen(js_name = "setUnlocked" , skip_jsdoc)]
    pub fn set_unlocked(&self) -> Format {
        impl_method!(self . set_unlocked ( ))
    }

    /// Unset the bold Format property back to its default "off" state.
    /// 
    /// The opposite of [`Format::setBold()`].
    #[wasm_bindgen(js_name = "unsetBold" , skip_jsdoc)]
    pub fn unset_bold(&self) -> Format {
        impl_method!(self . unset_bold ( ))
    }

    /// Unset the `checkbox` Format property back to its default "off" state.
    /// 
    /// The opposite of [`Format::setCheckbox()`].
    #[wasm_bindgen(js_name = "unsetCheckbox" , skip_jsdoc)]
    pub fn unset_checkbox(&self) -> Format {
        impl_method!(self . unset_checkbox ( ))
    }

    /// Unset the font strikethrough Format property back to its default "off" state.
    /// 
    /// The opposite of [`Format::setFontStrikethrough()`].
    #[wasm_bindgen(js_name = "unsetFontStrikethrough" , skip_jsdoc)]
    pub fn unset_font_strikethrough(&self) -> Format {
        impl_method!(self . unset_font_strikethrough ( ))
    }

    /// Unset the hidden Format property back to its default "off" state.
    /// 
    /// The opposite of [`Format::setHidden()`].
    #[wasm_bindgen(js_name = "unsetHidden" , skip_jsdoc)]
    pub fn unset_hidden(&self) -> Format {
        impl_method!(self . unset_hidden ( ))
    }

    /// Unset the hyperlink style. Doesn't reset the other properties.
    #[wasm_bindgen(js_name = "unsetHyperlinkStyle" , skip_jsdoc)]
    pub fn unset_hyperlink_style(&self) -> Format {
        impl_method!(self . unset_hyperlink_style ( ))
    }

    /// Unset the italic Format property back to its default "off" state.
    /// 
    /// The opposite of [`Format::setItalic()`].
    #[wasm_bindgen(js_name = "unsetItalic" , skip_jsdoc)]
    pub fn unset_italic(&self) -> Format {
        impl_method!(self . unset_italic ( ))
    }

    /// Unset the `quote_prefix` Format property back to its default "off" state.
    /// 
    /// The opposite of [`Format::setQuotePrefix()`].
    #[wasm_bindgen(js_name = "unsetQuotePrefix" , skip_jsdoc)]
    pub fn unset_quote_prefix(&self) -> Format {
        impl_method!(self . unset_quote_prefix ( ))
    }

    /// Unset the shrink Format property back to its default "off" state.
    /// 
    /// The opposite of [`Format::setShrink()`].
    #[wasm_bindgen(js_name = "unsetShrink" , skip_jsdoc)]
    pub fn unset_shrink(&self) -> Format {
        impl_method!(self . unset_shrink ( ))
    }

    /// Unset the text wrap Format property back to its default "off" state.
    /// 
    /// The opposite of [`Format::setTextWrap()`].
    #[wasm_bindgen(js_name = "unsetTextWrap" , skip_jsdoc)]
    pub fn unset_text_wrap(&self) -> Format {
        impl_method!(self . unset_text_wrap ( ))
    }
}

