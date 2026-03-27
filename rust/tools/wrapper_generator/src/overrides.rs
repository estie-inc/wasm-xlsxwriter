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
    rename: HashMap<String, String>,
    #[serde(default)]
    skip_enums: HashMap<String, String>,
    /// Override specific parameter types: "Struct::method::param" = "ParamType"
    #[serde(default)]
    param_type: HashMap<String, String>,
    /// Struct names whose auto-generated `new()` constructor should be suppressed
    /// (used when the companion file provides a custom constructor)
    #[serde(default)]
    skip_constructors: HashMap<String, String>,
    /// Explicit consume_self_default expressions for types without Default
    /// e.g. "Shape" = "xlsx::Shape::textbox()"
    #[serde(default)]
    consume_self_default: HashMap<String, String>,
}

/// Parsed result of overrides.toml. Holds per-method override information.
pub struct Overrides {
    /// Map of `"Struct::method"` to MethodOverride (wildcards are normalized to `"*::method"`)
    entries: HashMap<String, MethodOverride>,
    /// Enum names to skip generation (hand-written separately)
    pub skip_enums: HashSet<String>,
    /// Map of `"Struct::method::param"` to param type override string
    param_types: HashMap<String, String>,
    /// Struct names whose auto-generated constructor should be suppressed
    pub skip_constructors: HashSet<String>,
    /// Explicit consume_self_default expressions: "StructName" -> "xlsx::Struct::factory()"
    pub consume_self_default: HashMap<String, String>,
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
            skip_enums: raw.skip_enums.into_keys().collect(),
            param_types: raw.param_type,
            skip_constructors: raw.skip_constructors.into_keys().collect(),
            consume_self_default: raw.consume_self_default,
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

    pub fn should_skip_enum(&self, name: &str) -> bool {
        self.skip_enums.contains(name)
    }

    /// Get the param type override for a specific struct::method::param, if any.
    pub fn get_param_type(&self, struct_name: &str, method_name: &str, param_name: &str) -> Option<&str> {
        let key = format!("{}::{}::{}", struct_name, method_name, param_name);
        self.param_types.get(&key).map(|s| s.as_str())
    }

    /// Apply overrides to a list of methods.
    /// Only methods with `override_` set to `Auto` are eligible for overwriting.
    /// Also applies param_type overrides to individual parameters.
    pub fn apply_to_methods(&self, struct_name: &str, methods: &mut Vec<AnalyzedMethod>) {
        for method in methods.iter_mut() {
            if let MethodOverride::Auto = method.override_ {
                if let Some(ov) = self.get(struct_name, &method.name) {
                    method.override_ = ov;
                }
            }
            // Apply param type overrides
            for param in &mut method.params {
                if let Some(ty_str) = self.get_param_type(struct_name, &method.name, &param.name) {
                    param.ty = crate::ir::ParamType::from_override_str(ty_str);
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
"Workbook::deep_clone" = "clone"
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
    fn parse_rename_entries() {
        let ov = load_overrides_from_str(SAMPLE_TOML).unwrap();

        assert_eq!(
            ov.get("Workbook", "deep_clone"),
            Some(MethodOverride::Rename("clone".into()))
        );
        // Exact match only
        assert_eq!(ov.get("Chart", "deep_clone"), None);
    }

    #[test]
    fn get_returns_none_for_unlisted_methods() {
        let ov = load_overrides_from_str(SAMPLE_TOML).unwrap();

        assert_eq!(ov.get("Workbook", "unknown_method"), None);
        assert_eq!(ov.get("Worksheet", "save"), None);
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
        let toml = r#"
[rename]
"*.some_method" = "renamed"
"#;
        let ov = load_overrides_from_str(toml).unwrap();

        let mut methods = vec![make_method("some_method"), make_method("set_color")];

        ov.apply_to_methods("AnyStruct", &mut methods);

        assert_eq!(
            methods[0].override_,
            MethodOverride::Rename("renamed".into())
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

    #[test]
    fn parse_param_type_entries() {
        let toml = r#"
[param_type]
"ChartSeries::set_name::name" = "Str"
"#;
        let ov = load_overrides_from_str(toml).unwrap();
        assert_eq!(ov.get_param_type("ChartSeries", "set_name", "name"), Some("Str"));
        assert_eq!(ov.get_param_type("ChartSeries", "set_name", "other"), None);
        assert_eq!(ov.get_param_type("Other", "set_name", "name"), None);
    }

    #[test]
    fn apply_to_methods_applies_param_type_override() {
        use crate::ir::{AnalyzedParam, ParamType};

        let toml = r#"
[param_type]
"Chart::set_name::name" = "Str"
"#;
        let ov = load_overrides_from_str(toml).unwrap();

        let mut methods = vec![AnalyzedMethod {
            name: "set_name".into(),
            js_name: "setName".into(),
            receiver: crate::ir::ReceiverKind::MutSelf,
            params: vec![AnalyzedParam {
                name: "name".into(),
                ty: ParamType::RefWrappedType("ChartRange".into()),
            }],
            returns: crate::ir::ReturnKind::SelfType,
            override_: MethodOverride::Auto,
            doc: None,
        }];

        ov.apply_to_methods("Chart", &mut methods);

        assert_eq!(methods[0].params[0].ty, ParamType::Str);
    }

}
