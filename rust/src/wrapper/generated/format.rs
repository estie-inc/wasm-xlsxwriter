use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `Format` struct is used to define cell formatting for data in a
/// worksheet.
///
/// The properties of a cell that can be formatted include: fonts, colors,
/// patterns, borders, alignment and number formatting.
#[derive(Clone)]
#[wasm_bindgen]
pub struct Format {
    pub(crate) inner: Arc<Mutex<xlsx::Format>>,
}

#[wasm_bindgen]
impl Format {
    #[wasm_bindgen(constructor)]
    pub fn new() -> Format {
        Format {
            inner: Arc::new(Mutex::new(xlsx::Format::new())),
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
        let lock = self.inner.lock().unwrap();
        lock.merge(&other.inner);
        Format {
            inner: Arc::clone(&self.inner),
        }
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
    ///     crate::Format#number-formats-in-different-locales
    ///
    /// # Parameters
    ///
    /// - `num_format`: The number format property.
    #[wasm_bindgen(js_name = "setNumFormat", skip_jsdoc)]
    pub fn set_num_format(&self, num_format: &str) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_num_format(num_format);
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the number format for a Format using a legacy format index.
    ///
    /// This method is similar to {@link Format#setNumFormat} except that it
    /// uses an index to a limited number of Excel's built-in, and legacy,
    /// number formats.
    ///
    /// Unless you need to specifically access one of Excel's built-in number
    /// formats the {@link Format#setNumFormat} method is a better solution.
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
    ///  - Numeric formats 23 to 36 are not documented by Microsoft and may
    ///    differ in international versions. The listed date and currency
    ///    formats may also vary depending on system settings.
    ///  - The dollar sign in the above format appears as the defined local
    ///    currency symbol.
    ///  - These formats can also be set via
    ///    {@link Format#setNumFormat}.
    ///
    /// # Parameters
    ///
    /// - `num_format_index`: The index to one of the inbuilt formats shown in
    ///   the table above.
    #[wasm_bindgen(js_name = "setNumFormatIndex", skip_jsdoc)]
    pub fn set_num_format_index(&self, num_format_index: u8) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_num_format_index(num_format_index);
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the bold property for a Format font.
    #[wasm_bindgen(js_name = "setBold", skip_jsdoc)]
    pub fn set_bold(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_bold();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the italic property for the Format font.
    #[wasm_bindgen(js_name = "setItalic", skip_jsdoc)]
    pub fn set_italic(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_italic();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the Format font name property.
    ///
    /// Set the font for a cell format.
    ///
    /// # Parameters
    ///
    /// - `font_name`: The font name property.
    #[wasm_bindgen(js_name = "setFontName", skip_jsdoc)]
    pub fn set_font_name(&self, font_name: &str) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_font_name(font_name);
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the Format font size property.
    ///
    /// Set the font size of the cell format. The size is generally an integer
    /// value but Excel allows x.5 values (hence the property is a f64 or
    /// types that can convert `Into` a f64).
    ///
    /// Excel adjusts the height of a row to accommodate the largest font size
    /// in the row.
    ///
    /// # Parameters
    ///
    /// - `font_size`: The font size property.
    #[wasm_bindgen(js_name = "setFontSize", skip_jsdoc)]
    pub fn set_font_size(&self, font_size: f64) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_font_size(font_size);
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
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
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_font_family(font_family);
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
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
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_font_charset(font_charset);
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the Format font strikethrough property.
    #[wasm_bindgen(js_name = "setFontStrikethrough", skip_jsdoc)]
    pub fn set_font_strikethrough(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_font_strikethrough();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the Format text wrap property.
    ///
    /// This method is used to turn on automatic text wrapping for text in a
    /// cell. If you wish to control where the string is wrapped you can add
    /// newlines to the text (see the example below).
    ///
    /// Excel generally adjusts the height of the cell to fit the wrapped text
    /// unless a explicit row height has be set via
    /// {@link Worksheet#setRowHeight}).
    #[wasm_bindgen(js_name = "setTextWrap", skip_jsdoc)]
    pub fn set_text_wrap(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_text_wrap();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the Format indent property.
    ///
    /// This method can be used to indent text in a cell.
    ///
    /// Indentation is a horizontal alignment property. It can be used in Excel
    /// in conjunction with the {@link FormatAlign#Left}, {@link FormatAlign#Right}
    /// and {@link FormatAlign#Distributed} alignments. It will override any other
    /// horizontal properties that don't support indentation.
    ///
    /// # Parameters
    ///
    /// - `indent`: The indentation level for the cell.
    #[wasm_bindgen(js_name = "setIndent", skip_jsdoc)]
    pub fn set_indent(&self, indent: u8) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_indent(indent);
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
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
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_rotation(rotation);
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the Format text reading order property.
    ///
    /// Set the text reading direction. This is useful when creating Arabic,
    /// Hebrew or other near or far eastern worksheets. It can be used in
    /// conjunction with the Worksheet
    /// {@link set_right_to_left}) method
    /// which changes the cell display direction of the worksheet.
    ///
    /// # Parameters
    ///
    /// - `reading_direction`: The reading order property, should be 0, 1, or
    ///   2, where these values refer to:
    ///
    ///   0. The reading direction is determined heuristically by Excel
    ///      depending on the text. This is the default option.
    ///   1. The text is displayed Left-to-Right, like English.
    ///   2. The text is displayed Right-to-Left, like Hebrew or Arabic.
    #[wasm_bindgen(js_name = "setReadingDirection", skip_jsdoc)]
    pub fn set_reading_direction(&self, reading_direction: u8) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_reading_direction(reading_direction);
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the Format shrink property.
    ///
    /// This method can be used to shrink text so that it fits in a cell
    #[wasm_bindgen(js_name = "setShrink", skip_jsdoc)]
    pub fn set_shrink(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_shrink();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the hyperlink style.
    ///
    /// Set the hyperlink style for use with urls. This is usually set
    /// automatically when writing urls without a format applied.
    #[wasm_bindgen(js_name = "setHyperlink", skip_jsdoc)]
    pub fn set_hyperlink(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_hyperlink();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the Format cell unlocked state.
    ///
    /// This method can be used to allow modification of a cell in a protected
    /// worksheet. In Excel, cell locking is turned on by default for all cells.
    /// However, it only has an effect if the worksheet has been protected using
    /// the {@link Worksheet#protect} method.
    #[wasm_bindgen(js_name = "setUnlocked", skip_jsdoc)]
    pub fn set_unlocked(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_unlocked();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the Format property to hide formulas in a cell.
    ///
    /// This method can be used to hide a formula while still displaying its
    /// result. This is generally used to hide complex calculations from end
    /// users who are only interested in the result. It only has an effect if
    /// the worksheet has been protected using the
    /// {@link Worksheet#protect} method.
    ///
    /// See the example above.
    #[wasm_bindgen(js_name = "setHidden", skip_jsdoc)]
    pub fn set_hidden(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_hidden();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the `quote_prefix` property for a Format.
    ///
    /// Set the quote prefix property of a format to ensure a string is treated
    /// as a string after editing. This is the same as prefixing the string with
    /// a single quote in Excel. You don't need to add the quote to the string
    /// but you do need to add the format.
    #[wasm_bindgen(js_name = "setQuotePrefix", skip_jsdoc)]
    pub fn set_quote_prefix(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_quote_prefix();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the Format property to show a checkbox in a cell.
    ///
    /// This format property can be used with a cell that contains a boolean
    /// value to display it as a checkbox. This property isn't required very
    /// often and it is generally easier to create a checkbox using the
    /// {@link Worksheet#insertCheckbox}
    /// method.
    #[wasm_bindgen(js_name = "setCheckbox", skip_jsdoc)]
    pub fn set_checkbox(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_checkbox();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Unset the bold Format property back to its default "off" state.
    ///
    /// The opposite of {@link Format#setBold}.
    #[wasm_bindgen(js_name = "unsetBold", skip_jsdoc)]
    pub fn unset_bold(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.unset_bold();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Unset the italic Format property back to its default "off" state.
    ///
    /// The opposite of {@link Format#setItalic}.
    #[wasm_bindgen(js_name = "unsetItalic", skip_jsdoc)]
    pub fn unset_italic(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.unset_italic();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Unset the font strikethrough Format property back to its default "off" state.
    ///
    /// The opposite of {@link Format#setFontStrikethrough}.
    #[wasm_bindgen(js_name = "unsetFontStrikethrough", skip_jsdoc)]
    pub fn unset_font_strikethrough(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.unset_font_strikethrough();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Unset the text wrap Format property back to its default "off" state.
    ///
    /// The opposite of {@link Format#setTextWrap}.
    #[wasm_bindgen(js_name = "unsetTextWrap", skip_jsdoc)]
    pub fn unset_text_wrap(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.unset_text_wrap();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Unset the shrink Format property back to its default "off" state.
    ///
    /// The opposite of {@link Format#setShrink}.
    #[wasm_bindgen(js_name = "unsetShrink", skip_jsdoc)]
    pub fn unset_shrink(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.unset_shrink();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the locked Format property back to its default "on" state.
    ///
    /// The opposite of {@link Format#setUnlocked}.
    #[wasm_bindgen(js_name = "setLocked", skip_jsdoc)]
    pub fn set_locked(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_locked();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Unset the hidden Format property back to its default "off" state.
    ///
    /// The opposite of {@link Format#setHidden}.
    #[wasm_bindgen(js_name = "unsetHidden", skip_jsdoc)]
    pub fn unset_hidden(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.unset_hidden();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Unset the hyperlink style. Doesn't reset the other properties.
    #[wasm_bindgen(js_name = "unsetHyperlinkStyle", skip_jsdoc)]
    pub fn unset_hyperlink_style(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.unset_hyperlink_style();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Unset the `quote_prefix` Format property back to its default "off" state.
    ///
    /// The opposite of {@link Format#setQuotePrefix}.
    #[wasm_bindgen(js_name = "unsetQuotePrefix", skip_jsdoc)]
    pub fn unset_quote_prefix(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.unset_quote_prefix();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Unset the `checkbox` Format property back to its default "off" state.
    ///
    /// The opposite of {@link Format#setCheckbox}.
    #[wasm_bindgen(js_name = "unsetCheckbox", skip_jsdoc)]
    pub fn unset_checkbox(&self) -> Format {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.unset_checkbox();
        *lock = inner;
        Format {
            inner: Arc::clone(&self.inner),
        }
    }
}
