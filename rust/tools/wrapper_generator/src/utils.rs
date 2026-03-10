// Utilities for camelCase conversion, doc comment processing, etc.

/// Converts snake_case to camelCase for JS method names.
pub fn to_camel_case(snake: &str) -> String {
    let mut result = String::new();
    let mut capitalize_next = false;

    for ch in snake.chars() {
        if ch == '_' {
            capitalize_next = true;
        } else if capitalize_next {
            result.extend(ch.to_uppercase());
            capitalize_next = false;
        } else {
            result.push(ch);
        }
    }

    result
}

/// Transforms upstream rustdoc comments for wasm-bindgen JSDoc.
///
/// - Strips everything from `# Example` or `# Examples` section (inclusive)
/// - Converts `[`Format::set_bold()`]` → `Format.setBold()`
/// - Converts `[`StructName`]` → `StructName`
pub fn process_doc_comment(doc: &str) -> String {
    if doc.is_empty() {
        return String::new();
    }

    let truncated = truncate_at_examples_section(doc);
    let converted = convert_rust_refs(truncated);
    converted
}

fn truncate_at_examples_section(doc: &str) -> &str {
    // Find the first occurrence of "# Example" or "# Examples" as a section header.
    // We match at the start of a line (or start of string) to avoid false positives.
    let mut search_start = 0;
    loop {
        let remaining = &doc[search_start..];
        let Some(pos) = remaining.find("# Example") else {
            break;
        };
        let abs_pos = search_start + pos;

        // Must be at the start of the text or preceded by a newline.
        let at_line_start = abs_pos == 0 || doc.as_bytes()[abs_pos - 1] == b'\n';
        if !at_line_start {
            search_start = abs_pos + 1;
            continue;
        }

        // The character after "# Example" must be end-of-string, 's', newline, or whitespace
        // to distinguish "# Examples" and "# Example\n" from "# ExampleFoo".
        let after = &doc[abs_pos + "# Example".len()..];
        let valid_suffix = after.is_empty()
            || after.starts_with('s')
            || after.starts_with('\n')
            || after.starts_with('\r')
            || after.starts_with(' ');

        if valid_suffix {
            return doc[..abs_pos].trim_end();
        }

        search_start = abs_pos + 1;
    }

    doc.trim_end()
}

fn convert_rust_refs(doc: &str) -> String {
    let mut result = String::with_capacity(doc.len());
    let mut remaining = doc;

    while let Some(open) = remaining.find("[`") {
        result.push_str(&remaining[..open]);
        let after_open = &remaining[open + 2..];

        let Some(close) = after_open.find("`]") else {
            // Malformed reference — emit as-is and stop processing.
            result.push_str(&remaining[open..]);
            return result;
        };

        let inner = &after_open[..close];
        result.push_str(&convert_single_ref(inner));
        remaining = &after_open[close + 2..];
    }

    result.push_str(remaining);
    result
}

/// Converts a single reference inner text (without the surrounding [`...`]).
///
/// `Format::set_bold()` → `Format.setBold()`
/// `StructName` → `StructName`
fn convert_single_ref(inner: &str) -> String {
    // Strip trailing `()` if present, remember it for later.
    let (base, trailing_parens) = if inner.ends_with("()") {
        (&inner[..inner.len() - 2], true)
    } else {
        (inner, false)
    };

    // Handle `Struct::method` pattern.
    if let Some(sep) = base.find("::") {
        let struct_name = &base[..sep];
        let method_name = &base[sep + 2..];
        let js_method = to_camel_case(method_name);
        if trailing_parens {
            return format!("{}.{}()", struct_name, js_method);
        } else {
            return format!("{}.{}", struct_name, js_method);
        }
    }

    // Plain struct/type reference — return as-is (no conversion needed for PascalCase).
    if trailing_parens {
        format!("{}()", base)
    } else {
        base.to_string()
    }
}

/// Converts PascalCase struct name to snake_case for filenames.
///
/// Digit sequences are treated as their own word segment.
/// `ConditionalFormat2ColorScale` → `conditional_format_2_color_scale`
pub fn to_snake_case_filename(name: &str) -> String {
    let mut result = String::new();
    let chars: Vec<char> = name.chars().collect();

    for (i, &ch) in chars.iter().enumerate() {
        let prev = if i > 0 { Some(chars[i - 1]) } else { None };

        if ch.is_uppercase() {
            // Insert underscore before uppercase if:
            // - Not the first character, AND
            // - Previous character was lowercase/digit (normal word boundary), OR
            // - Previous was uppercase but next is lowercase (e.g., "ABCDef" → "abc_def")
            let next = chars.get(i + 1).copied();
            let needs_underscore = prev.map_or(false, |p| {
                p.is_lowercase()
                    || p.is_ascii_digit()
                    || (p.is_uppercase() && next.map_or(false, |n| n.is_lowercase()))
            });
            if needs_underscore {
                result.push('_');
            }
            result.extend(ch.to_lowercase());
        } else if ch.is_ascii_digit() {
            // Insert underscore before digit if previous was a letter.
            if prev.map_or(false, |p| p.is_alphabetic()) {
                result.push('_');
            }
            result.push(ch);
        } else {
            // Lowercase letter: insert underscore if previous was a digit.
            if prev.map_or(false, |p| p.is_ascii_digit()) {
                result.push('_');
            }
            result.push(ch);
        }
    }

    result
}

#[cfg(test)]
mod tests {
    use super::*;
    use pretty_assertions::assert_eq;

    // ── to_camel_case ──────────────────────────────────────────────────────────

    #[test]
    fn camel_case_multi_segment() {
        assert_eq!(to_camel_case("set_font_name"), "setFontName");
    }

    #[test]
    fn camel_case_single_segment() {
        assert_eq!(to_camel_case("set_bold"), "setBold");
    }

    #[test]
    fn camel_case_digit_segment() {
        assert_eq!(to_camel_case("x2_axis"), "x2Axis");
    }

    #[test]
    fn camel_case_no_underscore() {
        assert_eq!(to_camel_case("name"), "name");
    }

    #[test]
    fn camel_case_alt_text() {
        assert_eq!(to_camel_case("set_alt_text"), "setAltText");
    }

    #[test]
    fn camel_case_new() {
        assert_eq!(to_camel_case("new"), "new");
    }

    #[test]
    fn camel_case_empty() {
        assert_eq!(to_camel_case(""), "");
    }

    // ── process_doc_comment ────────────────────────────────────────────────────

    #[test]
    fn doc_comment_empty() {
        assert_eq!(process_doc_comment(""), "");
    }

    #[test]
    fn doc_comment_no_examples_section() {
        let doc = "Sets the font name for a Format object.";
        assert_eq!(process_doc_comment(doc), "Sets the font name for a Format object.");
    }

    #[test]
    fn doc_comment_strips_example_section() {
        let doc = "Sets the bold property.\n\n# Example\n\n```rust\nlet f = Format::new();\n```";
        assert_eq!(process_doc_comment(doc), "Sets the bold property.");
    }

    #[test]
    fn doc_comment_strips_examples_section() {
        let doc = "Sets the bold property.\n\n# Examples\n\n```rust\nlet f = Format::new();\n```";
        assert_eq!(process_doc_comment(doc), "Sets the bold property.");
    }

    #[test]
    fn doc_comment_converts_method_ref() {
        let doc = "See [`Format::set_bold()`] for details.";
        assert_eq!(process_doc_comment(doc), "See Format.setBold() for details.");
    }

    #[test]
    fn doc_comment_converts_struct_ref() {
        let doc = "Returns a [`Format`] object.";
        assert_eq!(process_doc_comment(doc), "Returns a Format object.");
    }

    #[test]
    fn doc_comment_multiple_refs() {
        let doc = "Use [`Format::set_bold()`] or [`Format::set_italic()`].";
        assert_eq!(
            process_doc_comment(doc),
            "Use Format.setBold() or Format.setItalic()."
        );
    }

    #[test]
    fn doc_comment_strips_and_converts() {
        let doc = "Use [`Format::set_bold()`] for bold.\n\n# Examples\n\n```rust\nlet f = Format::new();\n```";
        assert_eq!(
            process_doc_comment(doc),
            "Use Format.setBold() for bold."
        );
    }

    #[test]
    fn doc_comment_does_not_strip_mid_line_example() {
        // "# Example" appearing in the middle of a line should not trigger truncation.
        let doc = "This is not a # Example section header.";
        assert_eq!(
            process_doc_comment(doc),
            "This is not a # Example section header."
        );
    }

    #[test]
    fn doc_comment_alt_text_ref() {
        let doc = "See [`Format::set_alt_text()`].";
        assert_eq!(process_doc_comment(doc), "See Format.setAltText().");
    }

    // ── to_snake_case_filename ─────────────────────────────────────────────────

    #[test]
    fn snake_case_filename_format_align() {
        assert_eq!(to_snake_case_filename("FormatAlign"), "format_align");
    }

    #[test]
    fn snake_case_filename_chart_data_label() {
        assert_eq!(to_snake_case_filename("ChartDataLabel"), "chart_data_label");
    }

    #[test]
    fn snake_case_filename_with_digit() {
        assert_eq!(
            to_snake_case_filename("ConditionalFormat2ColorScale"),
            "conditional_format_2_color_scale"
        );
    }

    #[test]
    fn snake_case_filename_single_word() {
        assert_eq!(to_snake_case_filename("Format"), "format");
    }

    #[test]
    fn snake_case_filename_empty() {
        assert_eq!(to_snake_case_filename(""), "");
    }
}
