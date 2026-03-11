// Parse overrides.toml and apply per-method overrides

use std::collections::{HashMap, HashSet};
use std::path::Path;

use anyhow::Context;
use serde::Deserialize;

use crate::ir::{AnalyzedMethod, MethodOverride};

/// Raw deserialization struct for overrides.toml.
/// Keys are in the format `"StructName::method_name"` or `"*.method_name"`, values are reason strings.
#[derive(Debug, Deserialize, Default)]
struct RawOverrides {
    #[serde(default)]
    skip: HashMap<String, String>,
    #[serde(default)]
    custom: HashMap<String, String>,
    #[serde(default)]
    consume_self_default: HashMap<String, String>,
    #[serde(default)]
    rename: HashMap<String, String>,
    #[serde(default)]
    skip_structs: HashMap<String, String>,
    #[serde(default)]
    skip_enums: HashMap<String, String>,
}

/// Parsed result of overrides.toml. Holds per-method override information.
pub struct Overrides {
    /// Map of `"Struct::method"` to MethodOverride (wildcards are normalized to `"*::method"`)
    entries: HashMap<String, MethodOverride>,
    /// Map of struct name to dummy constructor expression for ConsumeSelf methods
    consume_self_defaults: HashMap<String, String>,
    /// Struct names to skip entirely during generation
    skip_structs: HashSet<String>,
    /// Enum names to skip entirely during generation
    skip_enums: HashSet<String>,
}

pub fn load_overrides(path: &Path) -> anyhow::Result<Overrides> {
    let content = std::fs::read_to_string(path)
        .with_context(|| format!("Failed to read overrides.toml: {}", path.display()))?;
    load_overrides_from_str(&content)
}

pub fn load_overrides_from_str(content: &str) -> anyhow::Result<Overrides> {
    let raw: RawOverrides =
        toml::from_str(content).context("Failed to parse overrides.toml")?;
    Ok(Overrides::from_raw(raw))
}

impl Overrides {
    fn from_raw(raw: RawOverrides) -> Self {
        let mut entries = HashMap::new();

        for (key, reason) in raw.skip {
            entries.insert(normalize_key(&key), MethodOverride::Skip(reason));
        }
        for (key, reason) in raw.custom {
            entries.insert(normalize_key(&key), MethodOverride::Custom(reason));
        }
        for (key, new_name) in raw.rename {
            entries.insert(normalize_key(&key), MethodOverride::Rename(new_name));
        }

        Self {
            entries,
            consume_self_defaults: raw.consume_self_default,
            skip_structs: raw.skip_structs.into_keys().collect(),
            skip_enums: raw.skip_enums.into_keys().collect(),
        }
    }

    /// Return the override for the given struct::method.
    /// Also matches wildcard patterns `"*::method"`.
    pub fn get(&self, struct_name: &str, method_name: &str) -> Option<MethodOverride> {
        let exact_key = format!("{}::{}", struct_name, method_name);
        let wildcard_key = format!("*::{}", method_name);

        self.entries
            .get(&exact_key)
            .or_else(|| self.entries.get(&wildcard_key))
            .cloned()
    }

    /// Return the consume_self_default expression for the given struct, if any.
    pub fn get_consume_self_default(&self, struct_name: &str) -> Option<&str> {
        self.consume_self_defaults.get(struct_name).map(|s| s.as_str())
    }

    pub fn should_skip_struct(&self, name: &str) -> bool {
        self.skip_structs.contains(name)
    }

    pub fn should_skip_enum(&self, name: &str) -> bool {
        self.skip_enums.contains(name)
    }

    /// Apply overrides to a list of methods.
    /// Only methods with `override_` set to `Auto` are eligible for overwriting.
    pub fn apply_to_methods(&self, struct_name: &str, methods: &mut Vec<AnalyzedMethod>) {
        for method in methods.iter_mut() {
            if let MethodOverride::Auto = method.override_ {
                if let Some(ov) = self.get(struct_name, &method.name) {
                    method.override_ = ov;
                }
            }
        }
    }
}

/// Keep `"Struct::method"` as-is, normalize `"*.method"` to `"*::method"`.
fn normalize_key(key: &str) -> String {
    // Convert `"*.method_name"` format to `"*::method_name"`
    if let Some(method) = key.strip_prefix("*.") {
        return format!("*::{}", method);
    }
    key.to_string()
}

#[cfg(test)]
mod tests {
    use super::*;
    use pretty_assertions::assert_eq;

    fn make_method(name: &str) -> AnalyzedMethod {
        use crate::ir::{ReceiverKind, ReturnKind};
        AnalyzedMethod {
            name: name.to_string(),
            js_name: name.to_string(),
            receiver: ReceiverKind::MutSelf,
            params: vec![],
            returns: ReturnKind::Void,
            override_: MethodOverride::Auto,
            doc: None,
        }
    }

    const SAMPLE_TOML: &str = r#"
[skip]
"Workbook::save" = "File I/O, not supported in WASM"
"Chart::validate" = "internal validation only"

[custom]
"Worksheet::write" = "ExcelData polymorphism"
"Workbook::add_worksheet" = "index tracking"

[rename]
"*.deep_clone" = "clone"
"#;

    #[test]
    fn parse_skip_entries() {
        let ov = load_overrides_from_str(SAMPLE_TOML).unwrap();

        assert_eq!(
            ov.get("Workbook", "save"),
            Some(MethodOverride::Skip("File I/O, not supported in WASM".into()))
        );
        assert_eq!(
            ov.get("Chart", "validate"),
            Some(MethodOverride::Skip("internal validation only".into()))
        );
    }

    #[test]
    fn parse_custom_entries() {
        let ov = load_overrides_from_str(SAMPLE_TOML).unwrap();

        assert_eq!(
            ov.get("Worksheet", "write"),
            Some(MethodOverride::Custom("ExcelData polymorphism".into()))
        );
        assert_eq!(
            ov.get("Workbook", "add_worksheet"),
            Some(MethodOverride::Custom("index tracking".into()))
        );
    }

    #[test]
    fn parse_rename_entries_including_wildcard() {
        let ov = load_overrides_from_str(SAMPLE_TOML).unwrap();

        // Wildcard matches any struct
        assert_eq!(
            ov.get("Workbook", "deep_clone"),
            Some(MethodOverride::Rename("clone".into()))
        );
        assert_eq!(
            ov.get("Chart", "deep_clone"),
            Some(MethodOverride::Rename("clone".into()))
        );
    }

    #[test]
    fn get_returns_none_for_unlisted_methods() {
        let ov = load_overrides_from_str(SAMPLE_TOML).unwrap();

        assert_eq!(ov.get("Workbook", "unknown_method"), None);
        assert_eq!(ov.get("Worksheet", "save"), None);
    }

    #[test]
    fn wildcard_does_not_match_partial_method_name() {
        let ov = load_overrides_from_str(SAMPLE_TOML).unwrap();

        // The wildcard for "deep_clone" should not match "clone"
        assert_eq!(ov.get("Workbook", "clone"), None);
        assert_eq!(ov.get("Workbook", "deep"), None);
    }

    #[test]
    fn apply_to_methods_sets_overrides_on_matching_methods() {
        let ov = load_overrides_from_str(SAMPLE_TOML).unwrap();

        let mut methods = vec![
            make_method("save"),
            make_method("add_worksheet"),
            make_method("close"),
        ];

        ov.apply_to_methods("Workbook", &mut methods);

        assert_eq!(
            methods[0].override_,
            MethodOverride::Skip("File I/O, not supported in WASM".into())
        );
        assert_eq!(
            methods[1].override_,
            MethodOverride::Custom("index tracking".into())
        );
        assert_eq!(methods[2].override_, MethodOverride::Auto);
    }

    #[test]
    fn apply_to_methods_does_not_overwrite_existing_non_auto_override() {
        let ov = load_overrides_from_str(SAMPLE_TOML).unwrap();

        let mut methods = vec![make_method("save")];
        // Should not overwrite when a different override is already set
        methods[0].override_ = MethodOverride::Custom("already set".into());

        ov.apply_to_methods("Workbook", &mut methods);

        assert_eq!(
            methods[0].override_,
            MethodOverride::Custom("already set".into())
        );
    }

    #[test]
    fn apply_to_methods_handles_wildcard_rename() {
        let ov = load_overrides_from_str(SAMPLE_TOML).unwrap();

        let mut methods = vec![make_method("deep_clone"), make_method("set_color")];

        ov.apply_to_methods("AnyStruct", &mut methods);

        assert_eq!(
            methods[0].override_,
            MethodOverride::Rename("clone".into())
        );
        assert_eq!(methods[1].override_, MethodOverride::Auto);
    }

    #[test]
    fn empty_sections_are_ok() {
        let empty = r#"
[skip]
[custom]
[rename]
"#;
        let ov = load_overrides_from_str(empty).unwrap();
        assert_eq!(ov.get("Workbook", "save"), None);
    }

    #[test]
    fn missing_sections_are_ok() {
        let only_skip = r#"
[skip]
"Workbook::save" = "File I/O, not supported in WASM"
"#;
        let ov = load_overrides_from_str(only_skip).unwrap();
        assert_eq!(
            ov.get("Workbook", "save"),
            Some(MethodOverride::Skip("File I/O, not supported in WASM".into()))
        );
        assert_eq!(ov.get("Worksheet", "write"), None);
    }
}
