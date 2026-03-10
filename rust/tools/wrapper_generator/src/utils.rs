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
/// - Strips fenced code blocks and `<img>` tags
/// - Strips everything from `# Example` or `# Examples` section (inclusive)
/// - Converts `[`Format::set_bold()`]` → `{@link Format#setBold}`
/// - Converts `[`StructName`]` → `{@link StructName}`
/// - Rust-internal types (`Into`, `Clone`, etc.) are shown as plain text
/// - Strips "or a type that can convert [`Into`]" boilerplate
/// - Collapses excessive blank lines
pub fn process_doc_comment(doc: &str) -> String {
    if doc.is_empty() {
        return String::new();
    }

    let no_code = strip_code_blocks(doc);
    let no_img = strip_img_tags(&no_code);
    let truncated = truncate_at_examples_section(&no_img);
    // convert_rust_refs handles both [`ref`] and trailing (crate::...) patterns;
    // strip_crate_links catches remaining [text](crate::...) markdown links
    let converted = convert_rust_refs(truncated);
    let no_crate = strip_crate_links(&converted);
    let cleaned = strip_into_boilerplate(&no_crate);
    collapse_blank_lines(&cleaned)
}

/// Like `process_doc_comment`, but also truncates at the first `#` section header.
/// Intended for struct/enum-level docs where long tutorials should be omitted.
pub fn process_struct_doc_comment(doc: &str) -> String {
    if doc.is_empty() {
        return String::new();
    }

    let no_code = strip_code_blocks(doc);
    let no_img = strip_img_tags(&no_code);
    let truncated = truncate_at_first_section_header(&no_img);
    let converted = convert_rust_refs(truncated);
    let no_crate = strip_crate_links(&converted);
    let cleaned = strip_into_boilerplate(&no_crate);
    collapse_blank_lines(&cleaned)
}

fn strip_code_blocks(doc: &str) -> String {
    let mut result = String::new();
    let mut in_code_block = false;

    for line in doc.lines() {
        if line.trim_start().starts_with("```") {
            in_code_block = !in_code_block;
            continue;
        }
        if !in_code_block {
            result.push_str(line);
            result.push('\n');
        }
    }

    // Remove orphaned code-intro lines (e.g., "created with the following code:")
    let cleaned: Vec<&str> = result
        .lines()
        .filter(|line| {
            let lower = line.to_lowercase();
            let is_code_intro = lower.contains("following code")
                || lower.contains("code below")
                || lower.contains("code to generate");
            !is_code_intro
        })
        .collect();

    cleaned.join("\n")
}

fn strip_img_tags(doc: &str) -> String {
    doc.lines()
        .filter(|line| !line.trim().starts_with("<img"))
        .collect::<Vec<_>>()
        .join("\n")
}

/// Collapses consecutive blank lines into at most one.
fn collapse_blank_lines(doc: &str) -> String {
    let mut result = String::new();
    let mut prev_blank = false;

    for line in doc.lines() {
        if line.trim().is_empty() {
            if !prev_blank {
                result.push('\n');
            }
            prev_blank = true;
        } else {
            prev_blank = false;
            result.push_str(line);
            result.push('\n');
        }
    }

    result.trim_end().to_string()
}

/// Truncates at the first `# ` section header (any header).
fn truncate_at_first_section_header(doc: &str) -> &str {
    for (i, line) in doc.lines().enumerate() {
        if i > 0 && line.starts_with("# ") {
            // Find the byte offset of this line in the string
            let byte_offset: usize = doc.lines().take(i).map(|l| l.len() + 1).sum();
            return doc[..byte_offset].trim_end();
        }
    }
    doc.trim_end()
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

        // Strip trailing Rust rustdoc link targets:
        // `(crate::...)` inline links or `[ref_name]` reference-style links
        if remaining.starts_with('(') {
            if let Some(paren_close) = remaining.find(')') {
                let link_target = &remaining[1..paren_close];
                if link_target.starts_with("crate::") {
                    remaining = &remaining[paren_close + 1..];
                }
            }
        } else if remaining.starts_with('[') {
            if let Some(bracket_close) = remaining.find(']') {
                remaining = &remaining[bracket_close + 1..];
            }
        }
    }

    result.push_str(remaining);
    result
}

/// Strips `[text](crate::...)` Rust-internal markdown links, keeping only the text.
fn strip_crate_links(doc: &str) -> String {
    let mut result = String::with_capacity(doc.len());
    let mut remaining = doc;

    while let Some(open) = remaining.find("](crate::") {
        // Find the matching `[` for the display text
        let bracket_open = remaining[..open].rfind('[');
        if let Some(bo) = bracket_open {
            result.push_str(&remaining[..bo]);
            let display_text = &remaining[bo + 1..open];
            result.push_str(display_text);
            // Skip past the `](crate::...)`
            let after_paren = &remaining[open + 2..];
            if let Some(paren_close) = after_paren.find(')') {
                remaining = &after_paren[paren_close + 1..];
            } else {
                remaining = after_paren;
            }
        } else {
            result.push_str(&remaining[..open + 1]);
            remaining = &remaining[open + 1..];
        }
    }

    result.push_str(remaining);
    result
}

/// Converts a single reference inner text (without the surrounding [`...`]).
///
/// `Format::set_bold()` → `{@link Format#setBold}`
/// `StructName` → `{@link StructName}`
/// Rust-internal types are returned as plain code text.
fn convert_single_ref(inner: &str) -> String {
    let (base, _trailing_parens) = if inner.ends_with("()") {
        (&inner[..inner.len() - 2], true)
    } else {
        (inner, false)
    };

    // Handle `Struct::method` pattern.
    if let Some(sep) = base.find("::") {
        let struct_name = &base[..sep];
        let method_name = &base[sep + 2..];

        if is_rust_internal_type(struct_name) {
            return format!("`{}`", inner);
        }

        let js_method = to_camel_case(method_name);
        return format!("{{@link {}#{}}}", struct_name, js_method);
    }

    if is_rust_internal_type(base) {
        return format!("`{}`", inner);
    }

    format!("{{@link {}}}", base)
}

/// Rust-internal types that should not be wrapped in `{@link}` in JSDoc.
fn is_rust_internal_type(name: &str) -> bool {
    matches!(
        name,
        "Into"
            | "AsRef"
            | "From"
            | "Clone"
            | "Debug"
            | "Display"
            | "Default"
            | "String"
            | "Vec"
            | "Option"
            | "Result"
            | "Box"
            | "Arc"
            | "Mutex"
            | "Iterator"
            | "IntoIterator"
            | "Send"
            | "Sync"
            | "Sized"
            | "u8"
            | "u16"
            | "u32"
            | "u64"
            | "i8"
            | "i16"
            | "i32"
            | "i64"
            | "f32"
            | "f64"
            | "bool"
            | "str"
            | "usize"
            | "isize"
    )
}

/// Strips sentences about Rust's `Into<T>` trait, which are irrelevant to JS users.
fn strip_into_boilerplate(doc: &str) -> String {
    let mut lines: Vec<&str> = doc.lines().collect();

    // Remove lines that are entirely about Into<T> conversion boilerplate
    lines.retain(|line| {
        let lower = line.to_lowercase();
        // "color can be a ... that converts Into" / "or a type that implements Into"
        let is_into_boilerplate = (lower.contains("that convert")
            || lower.contains("that implements"))
            && (lower.contains("`into`") || lower.contains("into<"));
        !is_into_boilerplate
    });

    let result = lines.join("\n");
    // Trim trailing empty lines
    result.trim_end().to_string()
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
        assert_eq!(
            process_doc_comment(doc),
            "See {@link Format#setBold} for details."
        );
    }

    #[test]
    fn doc_comment_converts_struct_ref() {
        let doc = "Returns a [`Format`] object.";
        assert_eq!(
            process_doc_comment(doc),
            "Returns a {@link Format} object."
        );
    }

    #[test]
    fn doc_comment_multiple_refs() {
        let doc = "Use [`Format::set_bold()`] or [`Format::set_italic()`].";
        assert_eq!(
            process_doc_comment(doc),
            "Use {@link Format#setBold} or {@link Format#setItalic}."
        );
    }

    #[test]
    fn doc_comment_strips_and_converts() {
        let doc = "Use [`Format::set_bold()`] for bold.\n\n# Examples\n\n```rust\nlet f = Format::new();\n```";
        assert_eq!(
            process_doc_comment(doc),
            "Use {@link Format#setBold} for bold."
        );
    }

    #[test]
    fn doc_comment_does_not_strip_mid_line_example() {
        let doc = "This is not a # Example section header.";
        assert_eq!(
            process_doc_comment(doc),
            "This is not a # Example section header."
        );
    }

    #[test]
    fn doc_comment_alt_text_ref() {
        let doc = "See [`Format::set_alt_text()`].";
        assert_eq!(process_doc_comment(doc), "See {@link Format#setAltText}.");
    }

    #[test]
    fn doc_comment_rust_internal_type_not_linked() {
        // Rust-internal types become `code`, not {@link}
        let doc = "See [`Into`] and [`Vec`] for details.";
        let result = process_doc_comment(doc);
        assert!(result.contains("`Into`"), "Into should be plain code: {}", result);
        assert!(result.contains("`Vec`"), "Vec should be plain code: {}", result);
        assert!(!result.contains("{@link"), "No {{@link}} for Rust types: {}", result);
    }

    #[test]
    fn doc_comment_mixed_rust_and_wrapper_types() {
        let doc = "Use a [`Color`] value. See also [`Default`].";
        let result = process_doc_comment(doc);
        assert!(result.contains("{@link Color}"), "Color should be linked: {}", result);
        assert!(result.contains("`Default`"), "Default should be plain code: {}", result);
    }

    #[test]
    fn doc_comment_strips_into_boilerplate() {
        let doc = "Set the font color.\n\n`color` can be a type that converts [`Into`] a Color.";
        assert_eq!(process_doc_comment(doc), "Set the font color.");
    }

    #[test]
    fn doc_comment_strips_code_blocks() {
        let doc = "Set the bold property.\n\n```rust\nlet f = Format::new().set_bold();\n```\n\nMore text.";
        assert_eq!(
            process_doc_comment(doc),
            "Set the bold property.\n\nMore text."
        );
    }

    #[test]
    fn doc_comment_strips_img_tags() {
        let doc = "Set the bold property.\n\n<img src=\"https://example.com/image.png\">\n\nMore text.";
        assert_eq!(
            process_doc_comment(doc),
            "Set the bold property.\n\nMore text."
        );
    }

    #[test]
    fn doc_comment_strips_code_intro_line() {
        let doc = "Description.\n\nThe output was created with the following code:\n\n```rust\ncode();\n```";
        assert_eq!(process_doc_comment(doc), "Description.");
    }

    #[test]
    fn struct_doc_truncates_at_first_header() {
        let doc = "A brief description.\n\nMore detail.\n\n# Contents\n\n- item 1\n\n# Section\n\nLong tutorial.";
        assert_eq!(
            process_struct_doc_comment(doc),
            "A brief description.\n\nMore detail."
        );
    }

    #[test]
    fn struct_doc_no_header_keeps_all() {
        let doc = "Just a description with no headers.";
        assert_eq!(
            process_struct_doc_comment(doc),
            "Just a description with no headers."
        );
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
