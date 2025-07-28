use regex::Regex;
use ruast::Attribute;

/// FIXME: Hacky workaround to pretty format source code
pub(crate) fn new_line() -> Attribute {
    Attribute::doc_comment("")
}

/// Add "/// " to the start of each line
fn add_doc_comment_marker(s: &str) -> String {
    s.lines()
        .map(|line| format!("/// {}", line.trim()))
        .collect::<Vec<_>>()
        .join("\n")
}

/// Find the first occurrence of "# Example" and return everything before it,
/// because it's not relevant for the generated javascript code
fn omit_after_example(s: &str) -> &str {
    if let Some(pos) = s.find("# Example") {
        s[..pos].trim()
    } else {
        s.trim()
    }
}

/// Convert a snake_case string to camelCase
pub(crate) fn to_camel_case(s: &str) -> String {
    let mut result = String::new();
    let mut capitalize_next = false;

    for c in s.chars() {
        if c == '_' {
            capitalize_next = true;
        } else if capitalize_next {
            result.push(c.to_ascii_uppercase());
            capitalize_next = false;
        } else {
            result.push(c);
        }
    }

    result
}

/// Convert method names inside square brackets to camelCase
///
/// For example, [`Format::set_italic()`] becomes [`Format::setItalic()`]
fn convert_bracket_methods_to_camel_case(s: &str) -> String {
    let re = Regex::new(r"\[\`([^`]+)::([\w_]+)(\([^)]*\))\`\]").unwrap();

    re.replace_all(s, |caps: &regex::Captures| {
        let module = &caps[1];
        let method = &caps[2];
        let args = &caps[3];

        format!("[`{}::{}{}`]", module, to_camel_case(method), args)
    })
    .to_string()
}

/// Process documentation comments by:
/// 1. Removing content after "# Example"
/// 2. Converting method names in square brackets to camelCase
/// 3. Adding doc comment markers ("/// ") to each line
///
/// This function combines the functionality of `omit_after_example`,
/// `convert_bracket_methods_to_camel_case`, and `add_doc_comment_marker`.
pub(crate) fn process_doc_comment(s: &str) -> String {
    let trimmed = omit_after_example(s);
    let camel_cased = convert_bracket_methods_to_camel_case(trimmed);
    add_doc_comment_marker(&camel_cased)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_convert_bracket_methods_to_camel_case() {
        let input = "This is a reference to [`Format::set_italic()`] method.";
        let expected = "This is a reference to [`Format::setItalic()`] method.";
        assert_eq!(convert_bracket_methods_to_camel_case(input), expected);

        // Test with multiple references
        let input = "See [`Format::set_italic()`] and [`Chart::add_series()`].";
        let expected = "See [`Format::setItalic()`] and [`Chart::addSeries()`].";
        assert_eq!(convert_bracket_methods_to_camel_case(input), expected);
    }

    #[test]
    fn test_process_doc_comment() {
        let input = "Unset the italic Format property back to its default \"off\" state.\n\nThe opposite of [`Format::set_italic()`].\n\n# Example\nThis should be removed.";
        let expected = "/// Unset the italic Format property back to its default \"off\" state.\n/// \n/// The opposite of [`Format::setItalic()`].";
        assert_eq!(process_doc_comment(input), expected);
    }
}
