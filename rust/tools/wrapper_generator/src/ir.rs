/// Intermediate Representation for wrapper generation.
///
/// An intermediate representation that transforms information parsed from
/// the upstream rustdoc-json (crate-inspector) into a form suitable for code generation.
/// Produced by the analyze module and consumed by the codegen module.

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
    /// Custom expression for `mem::replace` on types without Default that have ConsumeSelf methods
    pub consume_self_default: Option<String>,
    pub constructor: Option<AnalyzedConstructor>,
    pub methods: Vec<AnalyzedMethod>,
    pub doc: Option<String>,
    /// Whether the inner type implements Clone (enables deep_clone generation)
    pub has_clone: bool,
}

/// Whether a struct is a standalone type or a proxy owned by a parent.
#[derive(Debug, Clone, PartialEq)]
pub enum StructRole {
    /// A standalone type. Held directly via `Arc<Mutex<xlsx::T>>`.
    Standalone,
    /// A type owned by a parent. Accessed indirectly via `Arc<Mutex<xlsx::Parent>>` + accessor.
    /// When multiple parent methods return the same type (e.g., x_axis/y_axis), all are included in accessors.
    Proxy {
        parent_name: String,
        accessors: Vec<Accessor>,
    },
}

#[derive(Debug, Clone, PartialEq)]
pub struct Accessor {
    /// The upstream parent method name (e.g., "x_axis")
    pub parent_method: String,
    /// The JS-side name (e.g., "xAxis")
    pub js_name: String,
    /// For indexed accessors like `worksheet_from_index(usize)`.
    /// When set, the proxy stores a key field and passes it to the accessor.
    pub key_type: Option<ParamType>,
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

/// The receiver type of an upstream method. Determines the method body during code generation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ReceiverKind {
    /// `fn method(self, ...)` — lock → mem::take → call → replace
    ConsumeSelf,
    /// `fn method(&mut self, ...)` — lock → call directly
    MutSelf,
    /// `fn method(&self, ...)` — lock (shared) → call
    RefSelf,
}

/// The return type of an upstream method. Determines the wrapper's return type and Result wrapping.
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
    /// Any other return type not listed above (e.g., getters)
    Other(String),
}

/// Per-method override specified in overrides.toml
#[derive(Debug, Clone, PartialEq, Default)]
pub enum MethodOverride {
    /// Auto-generate (default)
    #[default]
    Auto,
    /// Skip generation (e.g., not supported in WASM)
    Skip(String),
    /// Use a hand-written implementation
    Custom(String),
    /// Override the JS name
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

/// The parameter type. Determines the conversion strategy during code generation.
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
    /// `&str` — wasm-bindgen automatically handles String <-> &str conversion
    Str,
    /// Another wasm-xlsxwriter wrapper type (e.g., "Color", "FormatAlign")
    /// Converted via `.into()` during codegen
    WrappedType(String),
    /// `Vec<T>`
    VecOf(Box<ParamType>),
    /// `Option<T>`
    OptionOf(Box<ParamType>),
    /// A reference to another wasm-xlsxwriter wrapper type (e.g., "&ChartFont")
    /// Converted via `&param.inner` during codegen
    RefWrappedType(String),
    /// A mutable reference to another wasm-xlsxwriter wrapper type (e.g., "&mut ChartFormat")
    /// Converted via `&mut param.inner` during codegen
    MutRefWrappedType(String),
    /// `Vec<T>` but the upstream type is `&[T]` (slice reference)
    RefSliceOf(Box<ParamType>),
    /// An unresolved type (requires custom handling)
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

/// The kind of an enum variant
#[derive(Debug, Clone, PartialEq)]
pub enum VariantKind {
    /// No data (e.g., `Left`)
    Plain,
    /// Tuple variant (e.g., `RGB(u32)`, `Theme(u8, u8)`)
    Tuple(Vec<String>),
    /// Struct variant (e.g., `Foo { x: u32, y: u32 }`)
    Struct(Vec<(String, String)>),
}

// ============================================================
// Impl blocks for convenience
// ============================================================

impl AnalyzedStruct {
    pub fn is_proxy(&self) -> bool {
        matches!(self.role, StructRole::Proxy { .. })
    }

    /// Returns only methods whose override_ is Auto
    pub fn auto_methods(&self) -> impl Iterator<Item = &AnalyzedMethod> {
        self.methods
            .iter()
            .filter(|m| matches!(m.override_, MethodOverride::Auto))
    }

    /// Returns only methods that can be auto-generated (excludes those with Unknown type
    /// parameters, ConsumeSelf patterns that cannot be safely generated, and methods
    /// returning references to other types which can't be wrapped in wasm_bindgen)
    pub fn generatable_methods(&self) -> impl Iterator<Item = &AnalyzedMethod> {
        let can_consume_self = self.has_default || self.consume_self_default.is_some();
        self.auto_methods()
            .filter(|m| m.params.iter().all(|p| !p.ty.has_unknown()))
            .filter(move |m| {
                if matches!(m.receiver, ReceiverKind::ConsumeSelf) {
                    if matches!(m.returns, ReturnKind::SelfType | ReturnKind::ResultSelf) {
                        return can_consume_self;
                    }
                    return false;
                }
                true
            })
            // Methods returning references to other types (e.g., &mut ChartAxis) can't
            // be safely returned from wasm_bindgen functions due to lifetime constraints
            .filter(|m| !matches!(&m.returns, ReturnKind::Other(s) if s.contains('&')))
    }
}

impl AnalyzedEnum {
    /// Returns whether the enum has any variants with data
    pub fn has_data_variants(&self) -> bool {
        self.variants
            .iter()
            .any(|v| !matches!(v.kind, VariantKind::Plain))
    }
}

impl ParamType {
    /// Parse a ParamType from override string (e.g., "Str", "Bool", "U32", "Color")
    pub fn from_override_str(s: &str) -> ParamType {
        match s {
            "Bool" => ParamType::Bool,
            "U8" => ParamType::U8,
            "U16" => ParamType::U16,
            "U32" => ParamType::U32,
            "U64" => ParamType::U64,
            "I8" => ParamType::I8,
            "I16" => ParamType::I16,
            "I32" => ParamType::I32,
            "I64" => ParamType::I64,
            "F32" => ParamType::F32,
            "F64" => ParamType::F64,
            "Usize" => ParamType::Usize,
            "Str" => ParamType::Str,
            other => ParamType::WrappedType(other.to_string()),
        }
    }

    /// The type name used in the wasm-bindgen signature
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
            ParamType::RefWrappedType(name) | ParamType::MutRefWrappedType(name) => name.clone(),
            ParamType::Unknown(s) => s.clone(),
        }
    }

    /// Recursively checks for Unknown types
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
            consume_self_default: None,
            constructor: None,
            methods: vec![],
            doc: None,
            has_clone: true,
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
                        key_type: None,
                    },
                    Accessor {
                        parent_method: "y_axis".into(),
                        js_name: "yAxis".into(),
                        key_type: None,
                    },
                ],
            },
            has_default: false,
            consume_self_default: None,
            constructor: None,
            methods: vec![],
            doc: None,
            has_clone: true,
        };
        assert!(s.is_proxy());
    }

    #[test]
    fn auto_methods_filters_overrides() {
        let s = AnalyzedStruct {
            name: "Format".into(),
            role: StructRole::Standalone,
            has_default: true,
            consume_self_default: None,
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
            has_clone: true,
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

    #[test]
    fn param_type_from_override_str() {
        assert_eq!(ParamType::from_override_str("Str"), ParamType::Str);
        assert_eq!(ParamType::from_override_str("Bool"), ParamType::Bool);
        assert_eq!(ParamType::from_override_str("U32"), ParamType::U32);
        assert_eq!(ParamType::from_override_str("F64"), ParamType::F64);
        assert_eq!(
            ParamType::from_override_str("Color"),
            ParamType::WrappedType("Color".into())
        );
    }
}
