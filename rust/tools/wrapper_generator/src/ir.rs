/// Intermediate Representation for wrapper generation.
///
/// upstream の rustdoc-json (crate-inspector) から解析した情報を、
/// コード生成に必要な形に変換した中間表現。
/// analyze モジュールが生成し、codegen モジュールが消費する。

#[derive(Debug, Clone, PartialEq)]
pub struct AnalyzedCrate {
    pub structs: Vec<AnalyzedStruct>,
    pub enums: Vec<AnalyzedEnum>,
}

// ============================================================
// Struct
// ============================================================

#[derive(Debug, Clone, PartialEq)]
pub struct AnalyzedStruct {
    pub name: String,
    pub role: StructRole,
    pub has_default: bool,
    pub constructor: Option<AnalyzedConstructor>,
    pub methods: Vec<AnalyzedMethod>,
    pub doc: Option<String>,
}

/// struct が独立した型か、親に所有された proxy か
#[derive(Debug, Clone, PartialEq)]
pub enum StructRole {
    /// 独立した型。`Arc<Mutex<xlsx::T>>` で直接保持。
    Standalone,
    /// 親に所有された型。`Arc<Mutex<xlsx::Parent>>` + accessor で間接アクセス。
    /// 同じ型を返す親メソッドが複数ある場合 (e.g., x_axis/y_axis) は accessors に全て含まれる。
    Proxy {
        parent_name: String,
        accessors: Vec<Accessor>,
    },
}

#[derive(Debug, Clone, PartialEq)]
pub struct Accessor {
    /// upstream の親メソッド名 (e.g., "x_axis")
    pub parent_method: String,
    /// JS 側の名前 (e.g., "xAxis")
    pub js_name: String,
}

#[derive(Debug, Clone, PartialEq)]
pub struct AnalyzedConstructor {
    pub params: Vec<AnalyzedParam>,
}

// ============================================================
// Method
// ============================================================

#[derive(Debug, Clone, PartialEq)]
pub struct AnalyzedMethod {
    pub name: String,
    pub js_name: String,
    pub receiver: ReceiverKind,
    pub params: Vec<AnalyzedParam>,
    pub returns: ReturnKind,
    pub override_: MethodOverride,
    pub doc: Option<String>,
}

/// upstream メソッドの receiver 型。コード生成時のメソッド本体を決定する。
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ReceiverKind {
    /// `fn method(self, ...)` — lock → mem::take → call → replace
    ConsumeSelf,
    /// `fn method(&mut self, ...)` — lock → 直接 call
    MutSelf,
    /// `fn method(&self, ...)` — lock (shared) → call
    RefSelf,
}

/// upstream メソッドの戻り型。ラッパーの戻り型と Result wrapping を決定する。
#[derive(Debug, Clone, PartialEq)]
pub enum ReturnKind {
    /// `-> Self` or `-> &mut Self` (builder chain)
    SelfType,
    /// `-> Result<Self, XlsxError>` or `-> Result<&mut Self, XlsxError>`
    ResultSelf,
    /// `-> Result<(), XlsxError>`
    ResultVoid,
    /// `-> ()` (no return)
    Void,
    /// getter 等、上記以外の戻り型
    Other(String),
}

/// overrides.toml で指定されたメソッド単位のオーバーライド
#[derive(Debug, Clone, PartialEq, Default)]
pub enum MethodOverride {
    /// 自動生成 (デフォルト)
    #[default]
    Auto,
    /// 生成しない (WASM 不可等)
    Skip(String),
    /// 手書き実装を使う
    Custom(String),
    /// JS 名を上書き
    Rename(String),
}

// ============================================================
// Param
// ============================================================

#[derive(Debug, Clone, PartialEq)]
pub struct AnalyzedParam {
    pub name: String,
    pub ty: ParamType,
}

/// パラメータの型。コード生成時の変換方法を決定する。
#[derive(Debug, Clone, PartialEq)]
pub enum ParamType {
    Bool,
    U8,
    U16,
    U32,
    U64,
    I8,
    I16,
    I32,
    I64,
    F32,
    F64,
    Usize,
    /// `&str` — wasm-bindgen が自動で String ↔ &str を処理
    Str,
    /// 他の wasm-xlsxwriter ラッパー型 (e.g., "Color", "FormatAlign")
    /// codegen 時に `.into()` で変換
    WrappedType(String),
    /// `Vec<T>`
    VecOf(Box<ParamType>),
    /// `Option<T>`
    OptionOf(Box<ParamType>),
    /// 他の wasm-xlsxwriter ラッパー型への参照 (e.g., "&ChartFont")
    /// codegen 時に `&param.into()` で変換
    RefWrappedType(String),
    /// `Vec<T>` だが upstream は `&[T]` (スライス参照)
    RefSliceOf(Box<ParamType>),
    /// 解決できなかった型 (custom 対応が必要)
    Unknown(String),
}

// ============================================================
// Enum
// ============================================================

#[derive(Debug, Clone, PartialEq)]
pub struct AnalyzedEnum {
    pub name: String,
    pub variants: Vec<AnalyzedVariant>,
    pub has_default: bool,
    pub doc: Option<String>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct AnalyzedVariant {
    pub name: String,
    pub kind: VariantKind,
    pub doc: Option<String>,
}

/// enum variant の種類
#[derive(Debug, Clone, PartialEq)]
pub enum VariantKind {
    /// データなし (e.g., `Left`)
    Plain,
    /// タプル variant (e.g., `RGB(u32)`, `Theme(u8, u8)`)
    Tuple(Vec<String>),
    /// 構造体 variant (e.g., `Foo { x: u32, y: u32 }`)
    Struct(Vec<(String, String)>),
}

// ============================================================
// Impl blocks for convenience
// ============================================================

impl AnalyzedStruct {
    pub fn is_proxy(&self) -> bool {
        matches!(self.role, StructRole::Proxy { .. })
    }

    /// override_ が Auto のメソッドのみ返す
    pub fn auto_methods(&self) -> impl Iterator<Item = &AnalyzedMethod> {
        self.methods
            .iter()
            .filter(|m| matches!(m.override_, MethodOverride::Auto))
    }

    /// 自動生成可能なメソッドのみ返す（Unknown 型パラメータを含むものを除外、
    /// 安全に生成できない ConsumeSelf パターンも除外）
    pub fn generatable_methods(&self) -> impl Iterator<Item = &AnalyzedMethod> {
        let has_default = self.has_default;
        self.auto_methods()
            .filter(|m| m.params.iter().all(|p| !p.ty.has_unknown()))
            .filter(move |m| {
                if matches!(m.receiver, ReceiverKind::ConsumeSelf) {
                    // ConsumeSelf + SelfType は mem::take で処理可能 (Default 必須)
                    if matches!(m.returns, ReturnKind::SelfType) {
                        return has_default;
                    }
                    // ConsumeSelf + 他の return type: 安全な生成パターンがない
                    return false;
                }
                true
            })
    }
}

impl AnalyzedEnum {
    /// データ付き variant があるか
    pub fn has_data_variants(&self) -> bool {
        self.variants
            .iter()
            .any(|v| !matches!(v.kind, VariantKind::Plain))
    }
}

impl ParamType {
    /// wasm-bindgen のシグネチャで使う型名
    pub fn to_rust_type_str(&self) -> String {
        match self {
            ParamType::Bool => "bool".into(),
            ParamType::U8 => "u8".into(),
            ParamType::U16 => "u16".into(),
            ParamType::U32 => "u32".into(),
            ParamType::U64 => "u64".into(),
            ParamType::I8 => "i8".into(),
            ParamType::I16 => "i16".into(),
            ParamType::I32 => "i32".into(),
            ParamType::I64 => "i64".into(),
            ParamType::F32 => "f32".into(),
            ParamType::F64 => "f64".into(),
            ParamType::Usize => "usize".into(),
            ParamType::Str => "&str".into(),
            ParamType::WrappedType(name) => name.clone(),
            ParamType::VecOf(inner) => format!("Vec<{}>", inner.to_rust_type_str()),
            ParamType::RefSliceOf(inner) => format!("Vec<{}>", inner.to_rust_type_str()),
            ParamType::OptionOf(inner) => format!("Option<{}>", inner.to_rust_type_str()),
            ParamType::RefWrappedType(name) => name.clone(),
            ParamType::Unknown(s) => s.clone(),
        }
    }

    /// Unknown 型を再帰的にチェック
    pub fn has_unknown(&self) -> bool {
        match self {
            ParamType::Unknown(_) => true,
            ParamType::VecOf(inner) | ParamType::RefSliceOf(inner) | ParamType::OptionOf(inner) => {
                inner.has_unknown()
            }
            _ => false,
        }
    }

    pub fn is_primitive(&self) -> bool {
        matches!(
            self,
            ParamType::Bool
                | ParamType::U8
                | ParamType::U16
                | ParamType::U32
                | ParamType::U64
                | ParamType::I8
                | ParamType::I16
                | ParamType::I32
                | ParamType::I64
                | ParamType::F32
                | ParamType::F64
                | ParamType::Usize
                | ParamType::Str
        )
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn struct_role_standalone() {
        let s = AnalyzedStruct {
            name: "Format".into(),
            role: StructRole::Standalone,
            has_default: true,
            constructor: None,
            methods: vec![],
            doc: None,
        };
        assert!(!s.is_proxy());
    }

    #[test]
    fn struct_role_proxy() {
        let s = AnalyzedStruct {
            name: "ChartAxis".into(),
            role: StructRole::Proxy {
                parent_name: "Chart".into(),
                accessors: vec![
                    Accessor {
                        parent_method: "x_axis".into(),
                        js_name: "xAxis".into(),
                    },
                    Accessor {
                        parent_method: "y_axis".into(),
                        js_name: "yAxis".into(),
                    },
                ],
            },
            has_default: false,
            constructor: None,
            methods: vec![],
            doc: None,
        };
        assert!(s.is_proxy());
    }

    #[test]
    fn auto_methods_filters_overrides() {
        let s = AnalyzedStruct {
            name: "Format".into(),
            role: StructRole::Standalone,
            has_default: true,
            constructor: None,
            methods: vec![
                AnalyzedMethod {
                    name: "set_bold".into(),
                    js_name: "setBold".into(),
                    receiver: ReceiverKind::ConsumeSelf,
                    params: vec![],
                    returns: ReturnKind::SelfType,
                    override_: MethodOverride::Auto,
                    doc: None,
                },
                AnalyzedMethod {
                    name: "validate".into(),
                    js_name: "validate".into(),
                    receiver: ReceiverKind::RefSelf,
                    params: vec![],
                    returns: ReturnKind::Void,
                    override_: MethodOverride::Skip("internal".into()),
                    doc: None,
                },
            ],
            doc: None,
        };
        let auto: Vec<_> = s.auto_methods().collect();
        assert_eq!(auto.len(), 1);
        assert_eq!(auto[0].name, "set_bold");
    }

    #[test]
    fn enum_has_data_variants() {
        let plain = AnalyzedEnum {
            name: "FormatAlign".into(),
            variants: vec![
                AnalyzedVariant {
                    name: "Left".into(),
                    kind: VariantKind::Plain,
                    doc: None,
                },
            ],
            has_default: false,
            doc: None,
        };
        assert!(!plain.has_data_variants());

        let with_data = AnalyzedEnum {
            name: "Color".into(),
            variants: vec![
                AnalyzedVariant {
                    name: "Red".into(),
                    kind: VariantKind::Plain,
                    doc: None,
                },
                AnalyzedVariant {
                    name: "RGB".into(),
                    kind: VariantKind::Tuple(vec!["u32".into()]),
                    doc: None,
                },
            ],
            has_default: false,
            doc: None,
        };
        assert!(with_data.has_data_variants());
    }

    #[test]
    fn param_type_to_rust_str() {
        assert_eq!(ParamType::Bool.to_rust_type_str(), "bool");
        assert_eq!(ParamType::Str.to_rust_type_str(), "&str");
        assert_eq!(
            ParamType::WrappedType("Color".into()).to_rust_type_str(),
            "Color"
        );
        assert_eq!(
            ParamType::VecOf(Box::new(ParamType::U8)).to_rust_type_str(),
            "Vec<u8>"
        );
    }
}
