// overrides.toml のパースとメソッド単位のオーバーライド適用

use std::collections::HashMap;
use std::path::Path;

use anyhow::Context;
use serde::Deserialize;

use crate::ir::{AnalyzedMethod, MethodOverride};

/// overrides.toml の生デシリアライズ構造体。
/// キーは `"StructName::method_name"` または `"*.method_name"` 形式、値は理由文字列。
#[derive(Debug, Deserialize, Default)]
struct RawOverrides {
    #[serde(default)]
    skip: HashMap<String, String>,
    #[serde(default)]
    custom: HashMap<String, String>,
    #[serde(default)]
    rename: HashMap<String, String>,
}

/// overrides.toml をパースした結果。メソッド単位のオーバーライド情報を保持する。
pub struct Overrides {
    /// `"Struct::method"` → MethodOverride のマップ（ワイルドカードは `"*::method"` に正規化）
    entries: HashMap<String, MethodOverride>,
}

pub fn load_overrides(path: &Path) -> anyhow::Result<Overrides> {
    let content = std::fs::read_to_string(path)
        .with_context(|| format!("overrides.toml の読み込みに失敗: {}", path.display()))?;
    load_overrides_from_str(&content)
}

pub fn load_overrides_from_str(content: &str) -> anyhow::Result<Overrides> {
    let raw: RawOverrides =
        toml::from_str(content).context("overrides.toml のパースに失敗")?;
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

        Self { entries }
    }

    /// 指定の struct::method に対するオーバーライドを返す。
    /// ワイルドカードパターン `"*::method"` も照合する。
    pub fn get(&self, struct_name: &str, method_name: &str) -> Option<MethodOverride> {
        let exact_key = format!("{}::{}", struct_name, method_name);
        let wildcard_key = format!("*::{}", method_name);

        self.entries
            .get(&exact_key)
            .or_else(|| self.entries.get(&wildcard_key))
            .cloned()
    }

    /// メソッドリストに対してオーバーライドを適用する。
    /// `override_` が `Auto` のメソッドのみ上書き対象とする。
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

/// `"Struct::method"` はそのまま、`"*.method"` は `"*::method"` に正規化する。
fn normalize_key(key: &str) -> String {
    // `"*.method_name"` 形式は `"*::method_name"` に変換する
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
"Workbook::save" = "File I/O、WASM不可"
"Chart::validate" = "内部検証用"

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
            Some(MethodOverride::Skip("File I/O、WASM不可".into()))
        );
        assert_eq!(
            ov.get("Chart", "validate"),
            Some(MethodOverride::Skip("内部検証用".into()))
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

        // ワイルドカードは任意の struct にマッチする
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

        // "deep_clone" のワイルドカードは "clone" にはマッチしない
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
            MethodOverride::Skip("File I/O、WASM不可".into())
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
        // 事前に別のオーバーライドが設定されている場合は上書きしない
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
"Workbook::save" = "File I/O、WASM不可"
"#;
        let ov = load_overrides_from_str(only_skip).unwrap();
        assert_eq!(
            ov.get("Workbook", "save"),
            Some(MethodOverride::Skip("File I/O、WASM不可".into()))
        );
        assert_eq!(ov.get("Worksheet", "write"), None);
    }
}
