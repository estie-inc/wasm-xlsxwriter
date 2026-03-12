use crate::wrapper::Color;
use crate::wrapper::Format;
use crate::wrapper::ObjectMovement;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `Note` struct represents a worksheet note object.
///
/// A Note is a post-it style message that is revealed when the user hovers over
/// a worksheet cell. The presence of a Note is indicated by a small red
/// triangle in the upper right-hand corner of the cell.
///
/// Notes are used in conjunction with the
/// {@link Worksheet#insertNote} method.
///
/// In versions of Excel prior to Office 365, Notes were referred to as
/// "Comments". The name "Comment" is now used for a newer style threaded comment,
/// and "Note" is used for the older non-threaded version. See the Microsoft docs
/// on [The difference between threaded comments and notes].
///
/// [The difference between threaded comments and notes]:
///     https://support.microsoft.com/en-us/office/the-difference-between-threaded-comments-and-notes-75a51eec-4092-42ab-abf8-7669077b7be3
///
/// Note, the newer Threaded Comments are unlikely to be added to
/// `rust_xlsxwriter` due to the fact that the feature relies on company
/// specific metadata to identify the comment author.
#[derive(Clone)]
#[wasm_bindgen]
pub struct Note {
    pub(crate) inner: Arc<Mutex<xlsx::Note>>,
}

#[wasm_bindgen]
impl Note {
    #[wasm_bindgen(constructor)]
    pub fn new(text: &str) -> Note {
        Note {
            inner: Arc::new(Mutex::new(xlsx::Note::new(text))),
        }
    }
    #[doc = r" Create a deep clone of this object."]
    #[wasm_bindgen(js_name = "clone")]
    pub fn deep_clone(&self) -> Note {
        Note {
            inner: Arc::new(Mutex::new(self.inner.lock().unwrap().clone())),
        }
    }
    /// Set the note author name.
    ///
    /// The author name appears in two places: at the start of the note text in
    /// bold and at the bottom of the worksheet in the status bar.
    ///
    /// If no name is specified, the default name "Author" will be applied to the
    /// note.
    ///
    /// You can also set the default author name for all notes in a worksheet
    /// via the
    /// {@link Worksheet#setDefaultNoteAuthor}
    /// method.
    ///
    /// # Parameters
    ///
    /// - `name`: The note author name. Must be less than or equal to the Excel
    ///   limit of 52 characters.
    #[wasm_bindgen(js_name = "setAuthor", skip_jsdoc)]
    pub fn set_author(&self, name: &str) -> Note {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Note::new(""));
        inner = inner.set_author(name);
        *lock = inner;
        Note {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Prefix the note text with the author name.
    ///
    /// By default Excel, and `rust_xlsxwriter`, prefixes the author name to the
    /// note text (see the previous examples). If you prefer to have the note
    /// text without the author name you can use this option to turn it off.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is on by default.
    #[wasm_bindgen(js_name = "addAuthorPrefix", skip_jsdoc)]
    pub fn add_author_prefix(&self, enable: bool) -> Note {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Note::new(""));
        inner = inner.add_author_prefix(enable);
        *lock = inner;
        Note {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Reset the text in the note.
    ///
    /// In general, the text of the note is set in the {@link Note#new}
    /// constructor, but if required, you can use the `reset_text()` method to
    /// reset the text for a note. This allows a single `Note` instance to be
    /// used multiple times and avoids the small overhead of creating a new
    /// instance each time.
    ///
    /// # Parameters
    ///
    /// - `text`: The text for the note.
    #[wasm_bindgen(js_name = "resetText", skip_jsdoc)]
    pub fn reset_text(&self, text: &str) -> Note {
        let mut lock = self.inner.lock().unwrap();
        lock.reset_text(text);
        Note {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the width of the note in pixels.
    ///
    /// The default width of an Excel note is 128 pixels.
    ///
    /// # Parameters
    ///
    /// - `width`: The note width in pixels.
    #[wasm_bindgen(js_name = "setWidth", skip_jsdoc)]
    pub fn set_width(&self, width: u32) -> Note {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Note::new(""));
        inner = inner.set_width(width);
        *lock = inner;
        Note {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the height of the note in pixels.
    ///
    /// The default height of an Excel note is 74 pixels. See the example above.
    ///
    /// # Parameters
    ///
    /// - `height`: The note height in pixels.
    #[wasm_bindgen(js_name = "setHeight", skip_jsdoc)]
    pub fn set_height(&self, height: u32) -> Note {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Note::new(""));
        inner = inner.set_height(height);
        *lock = inner;
        Note {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Make the note visible when the file loads.
    ///
    /// By default Excel hides cell notes until the user mouses over the parent
    /// cell. However, if required you can make the note visible without
    /// requiring an interaction from the user.
    ///
    /// You can also make all notes in a worksheet visible via the
    /// {@link Worksheet#showAllNotes}
    /// method.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setVisible", skip_jsdoc)]
    pub fn set_visible(&self, enable: bool) -> Note {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Note::new(""));
        inner = inner.set_visible(enable);
        *lock = inner;
        Note {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the background color for the note.
    ///
    /// The default background color for a Note is `#FFFFE1`. If required this
    /// method can be used to set it to a different RGB color.
    ///
    /// # Parameters
    ///
    /// - `color`: The background color property defined by a
    ///   {@link Color} enum value or a type that can convert `Into`
    ///   a {@link Color}. Only the `Color::Name` and `Color::RGB()` variants are
    ///   supported. Theme style colors aren't support by Excel for Notes.
    #[wasm_bindgen(js_name = "setBackgroundColor", skip_jsdoc)]
    pub fn set_background_color(&self, color: Color) -> Note {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Note::new(""));
        inner = inner.set_background_color(xlsx::Color::from(color));
        *lock = inner;
        Note {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the font name for the note.
    ///
    /// Set the font for a cell note. Excel can only display fonts that are
    /// installed on the system that it is running on. Therefore it is generally
    /// best to use standard Excel fonts.
    ///
    /// # Parameters
    ///
    /// - `font_name`: The font name for the note.
    #[wasm_bindgen(js_name = "setFontName", skip_jsdoc)]
    pub fn set_font_name(&self, font_name: &str) -> Note {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Note::new(""));
        inner = inner.set_font_name(font_name);
        *lock = inner;
        Note {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the font size for the note.
    ///
    /// Set the font size of the cell format. The size is generally an integer
    /// value but Excel allows x.5 values (hence the property is a f64 or
    /// types that can convert `Into` a f64).
    ///
    /// # Parameters
    ///
    /// - `font_size`: The font size for the note.
    #[wasm_bindgen(js_name = "setFontSize", skip_jsdoc)]
    pub fn set_font_size(&self, font_size: f64) -> Note {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Note::new(""));
        inner = inner.set_font_size(font_size);
        *lock = inner;
        Note {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the font family for the note.
    ///
    /// Set the font family. This is usually an integer in the range 1-4. This
    /// function is implemented for completeness but is rarely used in practice.
    ///
    /// # Parameters
    ///
    /// - `font_family`: The font family for the note.
    #[wasm_bindgen(js_name = "setFontFamily", skip_jsdoc)]
    pub fn set_font_family(&self, font_family: u8) -> Note {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Note::new(""));
        inner = inner.set_font_family(font_family);
        *lock = inner;
        Note {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the {@link Format} of the note.
    ///
    /// Set the font or background properties of a note using a {@link Format}
    /// object. Only the font name, size and background color are supported.
    ///
    /// This API is currently experimental and may go away in the future.
    ///
    /// # Parameters
    ///
    /// - `format`: The {@link Format} property for the note.
    #[wasm_bindgen(js_name = "setFormat", skip_jsdoc)]
    pub fn set_format(&self, format: Format) -> Note {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Note::new(""));
        inner = inner.set_format(format.inner.lock().unwrap().clone());
        *lock = inner;
        Note {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the alt text for the note to help accessibility.
    ///
    /// The alt text is used with screen readers to help people with visual
    /// disabilities.
    ///
    /// See the following Microsoft documentation on [Everything you need to
    /// know to write effective alt
    /// text](https://support.microsoft.com/en-us/office/everything-you-need-to-know-to-write-effective-alt-text-df98f884-ca3d-456c-807b-1a1fa82f5dc2).
    ///
    /// # Parameters
    ///
    /// - `alt_text`: The alt text string to add to the note.
    #[wasm_bindgen(js_name = "setAltText", skip_jsdoc)]
    pub fn set_alt_text(&self, alt_text: &str) -> Note {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Note::new(""));
        inner = inner.set_alt_text(alt_text);
        *lock = inner;
        Note {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the object movement options for a worksheet note.
    ///
    /// Set the option to define how an note will behave in Excel if the cells
    /// under the note are moved, deleted, or have their size changed. In
    /// Excel the options are:
    ///
    /// 1. Move and size with cells.
    /// 2. Move but don't size with cells.
    /// 3. Don't move or size with cells.
    ///
    /// These values are defined in the {@link ObjectMovement} enum.
    ///
    /// The {@link ObjectMovement} enum also provides an additional option to "Move
    /// and size with cells - after the note is inserted" to allow notes to
    /// be hidden in rows or columns. In Excel this equates to option 1 above
    /// but the internal note position calculations are handled differently.
    ///
    /// # Parameters
    ///
    /// - `option`: An note/object positioning behavior defined by the
    ///   {@link ObjectMovement} enum.
    #[wasm_bindgen(js_name = "setObjectMovement", skip_jsdoc)]
    pub fn set_object_movement(&self, option: ObjectMovement) -> Note {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Note::new(""));
        inner = inner.set_object_movement(xlsx::ObjectMovement::from(option));
        *lock = inner;
        Note {
            inner: Arc::clone(&self.inner),
        }
    }
}
