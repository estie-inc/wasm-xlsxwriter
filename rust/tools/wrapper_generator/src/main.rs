mod analyze;
mod codegen;
mod codegen_enum;
mod codegen_tokens;
mod ir;
mod overrides;
mod utils;

use std::path::{Path, PathBuf};

use anyhow::{Context, Result};
use clap::{Parser, Subcommand};

use crate::ir::{AnalyzedEnum, MethodOverride, StructRole, VariantKind};
use crate::utils::to_snake_case_filename;

#[derive(Parser)]
#[command(name = "wrapper_generator")]
#[command(about = "Auto-generate wasm-bindgen wrappers for rust_xlsxwriter")]
struct Cli {
    #[command(subcommand)]
    command: Command,
}

#[derive(Subcommand)]
enum Command {
    /// Generate wrapper code for all structs/enums
    Generate {
        /// Generate only specific types (comma-separated)
        #[arg(long)]
        filter: Option<String>,

        /// Output directory
        #[arg(long, default_value = "../../src/wrapper/generated")]
        output_dir: PathBuf,

        /// Path to overrides.toml
        #[arg(long, default_value = "overrides.toml")]
        overrides: PathBuf,

        /// Path to the upstream rust_xlsxwriter Cargo.toml (auto-resolved via cargo metadata if omitted)
        #[arg(long)]
        manifest: Option<PathBuf>,

        /// Workspace Cargo.toml used to auto-resolve the upstream crate (default: ../../Cargo.toml)
        #[arg(long, default_value = "../../Cargo.toml")]
        workspace_manifest: PathBuf,

        /// Name of the upstream crate to wrap (default: rust_xlsxwriter)
        #[arg(long, default_value = "rust_xlsxwriter")]
        upstream_crate: String,

        /// Directory of existing wrappers (types found here are excluded from mod.rs)
        #[arg(long)]
        existing_dir: Option<PathBuf>,
    },
    /// Display a summary of the upstream API analysis
    Verify {
        /// Path to overrides.toml
        #[arg(long, default_value = "overrides.toml")]
        overrides: PathBuf,

        /// Path to the upstream rust_xlsxwriter Cargo.toml (auto-resolved via cargo metadata if omitted)
        #[arg(long)]
        manifest: Option<PathBuf>,

        /// Workspace Cargo.toml used to auto-resolve the upstream crate (default: ../../Cargo.toml)
        #[arg(long, default_value = "../../Cargo.toml")]
        workspace_manifest: PathBuf,

        /// Name of the upstream crate to wrap (default: rust_xlsxwriter)
        #[arg(long, default_value = "rust_xlsxwriter")]
        upstream_crate: String,
    },
    /// Show differences between generated code and existing handwritten code
    Diff {
        /// Show only a specific struct
        #[arg(long, value_name = "NAME")]
        struct_name: Option<String>,

        /// Path to overrides.toml
        #[arg(long, default_value = "overrides.toml")]
        overrides: PathBuf,

        /// Path to the upstream rust_xlsxwriter Cargo.toml (auto-resolved via cargo metadata if omitted)
        #[arg(long)]
        manifest: Option<PathBuf>,

        /// Workspace Cargo.toml used to auto-resolve the upstream crate (default: ../../Cargo.toml)
        #[arg(long, default_value = "../../Cargo.toml")]
        workspace_manifest: PathBuf,

        /// Name of the upstream crate to wrap (default: rust_xlsxwriter)
        #[arg(long, default_value = "rust_xlsxwriter")]
        upstream_crate: String,

        /// Directory of existing wrappers
        #[arg(long, default_value = "../../src/wrapper")]
        existing_dir: PathBuf,
    },
}

fn main() -> Result<()> {
    let cli = Cli::parse();
    match cli.command {
        Command::Generate {
            filter,
            output_dir,
            overrides,
            manifest,
            workspace_manifest,
            upstream_crate,
            existing_dir,
        } => {
            let manifest = resolve_manifest(manifest, &workspace_manifest, &upstream_crate)?;
            cmd_generate(
                &manifest,
                &overrides,
                &output_dir,
                filter.as_deref(),
                existing_dir.as_deref(),
            )
        }
        Command::Verify {
            overrides,
            manifest,
            workspace_manifest,
            upstream_crate,
        } => {
            let manifest = resolve_manifest(manifest, &workspace_manifest, &upstream_crate)?;
            cmd_verify(&manifest, &overrides)
        }
        Command::Diff {
            struct_name,
            overrides,
            manifest,
            workspace_manifest,
            upstream_crate,
            existing_dir,
        } => {
            let manifest = resolve_manifest(manifest, &workspace_manifest, &upstream_crate)?;
            cmd_diff(
                &manifest,
                &overrides,
                &existing_dir,
                struct_name.as_deref(),
            )
        }
    }
}

/// If --manifest is provided, use it directly. Otherwise, auto-resolve via cargo metadata.
fn resolve_manifest(
    explicit: Option<PathBuf>,
    workspace_manifest: &Path,
    upstream_crate: &str,
) -> Result<PathBuf> {
    match explicit {
        Some(path) => Ok(path),
        None => {
            eprintln!(
                "Auto-resolving upstream crate '{}' via cargo metadata...",
                upstream_crate
            );
            resolve_upstream_manifest(workspace_manifest, upstream_crate)
        }
    }
}

struct AnalysisResult {
    analyzed: ir::AnalyzedCrate,
    overrides: overrides::Overrides,
    /// Enum names from the upstream crate before skip_enums filtering
    all_upstream_enum_names: std::collections::HashSet<String>,
}

fn load_and_analyze(
    manifest: &Path,
    overrides_path: &Path,
) -> Result<AnalysisResult> {
    eprintln!(
        "Loading upstream crate from: {}",
        manifest.display()
    );

    // Include #[doc(hidden)] items so we can wrap experimental/internal methods like Note::set_format
    std::env::set_var("RUSTDOCFLAGS", "--document-hidden-items");
    let krate = crate_inspector::CrateBuilder::default()
        .manifest_path(manifest)
        .document_private_items(true)
        .build()
        .context("Failed to generate rustdoc-json for the upstream crate")?;

    let mut analyzed = analyze::analyze_crate(&krate);

    let ov = overrides::load_overrides(overrides_path)?;

    // Capture all upstream enum names before filtering
    let all_upstream_enum_names: std::collections::HashSet<String> =
        analyzed.enums.iter().map(|e| e.name.clone()).collect();

    // Filter out enums listed in skip_enums (doc-hidden types are already excluded during analysis)
    analyzed.enums.retain(|e| !ov.should_skip_enum(&e.name));

    for s in &mut analyzed.structs {
        ov.apply_to_methods(&s.name, &mut s.methods);
    }

    // Re-evaluate proxy roles: if all parent methods creating the relationship are
    // *skipped* (not custom), the child should be standalone. E.g., ChartSeries is standalone
    // despite Chart::add_series returning &mut ChartSeries, because add_series is skipped.
    // Custom methods preserve the proxy relationship (the accessor is hand-written).
    for s in &mut analyzed.structs {
        if let StructRole::Proxy { parent_name, accessors } = &s.role {
            let all_skipped = accessors.iter().all(|acc| {
                ov.get(parent_name, &acc.parent_method)
                    .is_some_and(|o| matches!(o, MethodOverride::Skip(_)))
            });
            if all_skipped {
                s.role = StructRole::Standalone;
            }
        }
    }

    Ok(AnalysisResult {
        analyzed,
        overrides: ov,
        all_upstream_enum_names,
    })
}

/// Resolve the manifest path for an upstream dependency using cargo metadata.
fn resolve_upstream_manifest(workspace_manifest: &Path, crate_name: &str) -> Result<PathBuf> {
    let output = std::process::Command::new("cargo")
        .args(["metadata", "--format-version=1", "--manifest-path"])
        .arg(workspace_manifest)
        .output()
        .context("Failed to run cargo metadata")?;

    if !output.status.success() {
        let stderr = String::from_utf8_lossy(&output.stderr);
        anyhow::bail!("cargo metadata failed: {}", stderr);
    }

    let metadata: serde_json::Value = serde_json::from_slice(&output.stdout)
        .context("Failed to parse cargo metadata output")?;

    let packages = metadata["packages"]
        .as_array()
        .context("No packages in metadata")?;

    for pkg in packages {
        if pkg["name"].as_str() == Some(crate_name) {
            if let Some(path) = pkg["manifest_path"].as_str() {
                return Ok(PathBuf::from(path));
            }
        }
    }

    anyhow::bail!(
        "Crate '{}' not found in dependencies of {}",
        crate_name,
        workspace_manifest.display()
    )
}

// ── generate ────────────────────────────────────────────────────────────────

fn cmd_generate(
    manifest: &Path,
    overrides_path: &Path,
    output_dir: &Path,
    filter: Option<&str>,
    existing_dir: Option<&Path>,
) -> Result<()> {
    let AnalysisResult { analyzed, overrides: ov, all_upstream_enum_names } = load_and_analyze(manifest, overrides_path)?;
    let filter_set: Option<Vec<&str>> =
        filter.map(|f| f.split(',').map(str::trim).collect());

    // Collect type names from the existing directory (excluded from mod.rs to avoid duplicates)
    let existing = existing_dir
        .map(|dir| collect_existing_types(dir))
        .unwrap_or_else(|| ExistingTypes {
            all: Default::default(),
            arc_mutex_structs: Default::default(),
        });
    let existing_types = &existing.all;

    if !existing_types.is_empty() {
        eprintln!(
            "Excluding {} existing types from mod.rs",
            existing_types.len()
        );
    }

    // all_upstream_enum_names includes skip_enums (captured before filtering)
    let crate_enum_names = all_upstream_enum_names;
    let crate_struct_names: std::collections::HashSet<String> =
        analyzed.structs.iter().map(|s| s.name.clone()).collect();

    std::fs::create_dir_all(output_dir)
        .with_context(|| format!("Failed to create output directory: {}", output_dir.display()))?;
    let enums_dir = output_dir.join("enums");
    std::fs::create_dir_all(&enums_dir)?;

    let mut struct_modules = Vec::new();
    let mut enum_modules = Vec::new();
    let mut skipped_structs = Vec::new();
    let mut skipped_enums = Vec::new();

    // Process enums first to determine the set of available type names
    for e in &analyzed.enums {
        if let Some(ref f) = filter_set {
            if !f.contains(&e.name.as_str()) {
                continue;
            }
        }
        let code = codegen_enum::generate_enum_file(e);
        let module = to_snake_case_filename(&e.name);
        let filename = format!("{}.rs", module);
        let path = enums_dir.join(&filename);
        std::fs::write(&path, &code)?;

        if existing_types.contains(&e.name)
            || has_unresolvable_variants(e, &crate_enum_names, &crate_struct_names)
        {
            skipped_enums.push(e.name.clone());
        } else {
            enum_modules.push(module);
        }
        eprintln!("  enum:   {} -> {}", e.name, path.display());
    }

    // Collect all type names present in wrappers (existing + generated structs/enums - skipped)
    // Exclude unresolvable enums and skipped structs from available_types,
    // but keep skip_enums entries (they have hand-written .rs files)
    let skipped_set: std::collections::HashSet<&str> = skipped_enums
        .iter()
        .chain(skipped_structs.iter())
        .map(String::as_str)
        .collect();
    let mut available_types: std::collections::HashSet<String> = existing_types
        .iter()
        .cloned()
        .chain(crate_enum_names.iter().cloned())
        .chain(crate_struct_names.iter().cloned())
        .filter(|name| !skipped_set.contains(name.as_str()))
        .collect();
    // skip_enums entries are hand-written separately, so they're still available as types
    for name in &ov.skip_enums {
        available_types.insert(name.clone());
    }

    let codegen_ctx = codegen::CodegenContext {
        enum_names: crate_enum_names.clone(),
        available_types,
        // Only include hand-written structs that use direct `inner: xlsx::T` (not Arc<Mutex>)
        handwritten_struct_names: existing_types
            .iter()
            .filter(|name| crate_struct_names.contains(*name))
            .filter(|name| !existing.arc_mutex_structs.contains(*name))
            .cloned()
            .collect(),
        handwritten_enum_names: existing_types
            .iter()
            .filter(|name| crate_enum_names.contains(*name))
            .cloned()
            .collect(),
        skip_constructors: ov.skip_constructors.clone(),
    };

    for s in &analyzed.structs {
        if let Some(ref f) = filter_set {
            if !f.contains(&s.name.as_str()) {
                continue;
            }
        }
        let code = codegen::generate_struct_file(s, &codegen_ctx);
        let module = to_snake_case_filename(&s.name);
        let filename = format!("{}.rs", module);
        let path = output_dir.join(&filename);
        std::fs::write(&path, &code)?;

        if existing_types.contains(&s.name) {
            skipped_structs.push(s.name.clone());
        } else {
            struct_modules.push(module);
        }
        eprintln!("  struct: {} -> {}", s.name, path.display());
    }

    // Include skip_enums in enum mod.rs if they have a hand-written .rs file in the enums dir
    // (e.g. TableFunction) and aren't already provided by hand-written wrapper code outside generated/
    for name in &ov.skip_enums {
        if existing_types.contains(name) {
            continue;
        }
        let module = to_snake_case_filename(name);
        if enums_dir.join(format!("{module}.rs")).exists() {
            enum_modules.push(module);
        }
    }

    write_mod_file(output_dir, &struct_modules, true)?;
    write_mod_file(&enums_dir, &enum_modules, false)?;

    run_rustfmt(output_dir, &enums_dir);

    eprintln!(
        "\nDone: {} structs ({} skipped), {} enums ({} skipped).",
        struct_modules.len(),
        skipped_structs.len(),
        enum_modules.len(),
        skipped_enums.len(),
    );
    Ok(())
}

/// Scan results from existing wrapper directory.
struct ExistingTypes {
    /// All type names found (struct and enum)
    all: std::collections::HashSet<String>,
    /// Struct names that use `Arc<Mutex<>>` pattern (same as generated structs)
    arc_mutex_structs: std::collections::HashSet<String>,
}

/// Recursively scan .rs files in the existing wrapper directory and
/// collect type names from `pub struct TypeName` / `pub enum TypeName` declarations.
/// Also detect whether structs use `Arc<Mutex<>>` for their inner field.
fn collect_existing_types(dir: &Path) -> ExistingTypes {
    use std::collections::HashSet;
    let mut all = HashSet::new();
    let mut arc_mutex_structs = HashSet::new();

    if let Ok(entries) = walkdir(dir) {
        for path in entries {
            if path.extension().is_some_and(|ext| ext == "rs") {
                if let Ok(content) = std::fs::read_to_string(&path) {
                    let mut current_struct: Option<String> = None;
                    for line in content.lines() {
                        let trimmed = line.trim();

                        if let Some(rest) = trimmed.strip_prefix("pub struct ") {
                            let type_name: String =
                                rest.chars().take_while(|c| c.is_alphanumeric() || *c == '_').collect();
                            if !type_name.is_empty() {
                                all.insert(type_name.clone());
                                current_struct = Some(type_name);
                            }
                        } else if let Some(rest) = trimmed.strip_prefix("pub enum ") {
                            let type_name: String =
                                rest.chars().take_while(|c| c.is_alphanumeric() || *c == '_').collect();
                            if !type_name.is_empty() {
                                all.insert(type_name);
                            }
                            current_struct = None;
                        } else if let Some(ref name) = current_struct {
                            // Detect Arc<Mutex<>> pattern in inner field
                            if trimmed.contains("Arc<Mutex<") && trimmed.contains("inner") {
                                arc_mutex_structs.insert(name.clone());
                                current_struct = None;
                            } else if trimmed == "}" {
                                current_struct = None;
                            }
                        }
                    }
                }
            }
        }
    }

    ExistingTypes { all, arc_mutex_structs }
}

/// Simple recursive directory traversal (excludes generated/)
fn walkdir(dir: &Path) -> std::io::Result<Vec<PathBuf>> {
    let mut files = Vec::new();
    for entry in std::fs::read_dir(dir)? {
        let entry = entry?;
        let path = entry.path();
        if path.is_dir() {
            // Skip the generated/ directory itself
            if path.file_name().is_some_and(|n| n == "generated") {
                continue;
            }
            files.extend(walkdir(&path)?);
        } else {
            files.push(path);
        }
    }
    Ok(files)
}

/// Returns true if enum variants contain generics or types incompatible with tsify.
/// Struct types use Arc<Mutex> and are tsify/serde-incompatible. Enum types are tsify-compatible.
fn has_unresolvable_variants(
    e: &AnalyzedEnum,
    known_enum_names: &std::collections::HashSet<String>,
    known_struct_names: &std::collections::HashSet<String>,
) -> bool {
    let known_primitives = [
        "bool", "u8", "u16", "u32", "u64", "i8", "i16", "i32", "i64", "f32", "f64", "usize",
        "String",
    ];
    let is_unresolvable = |ty: &str| {
        if known_primitives.contains(&ty) {
            return false;
        }
        // Enums within the crate are tsify/serde-compatible
        if known_enum_names.contains(ty) {
            return false;
        }
        // Structs within the crate use Arc<Mutex> and are tsify-incompatible
        if known_struct_names.contains(ty) {
            return true;
        }
        // A single uppercase letter is a generic type parameter
        if ty.len() == 1 && ty.chars().next().is_some_and(|c| c.is_uppercase()) {
            return true;
        }
        // Types with paths or generic parameters
        if ty.contains('<') || ty.contains("::") {
            return true;
        }
        // Unknown PascalCase types (e.g., Error, ZipError)
        if ty.chars().next().is_some_and(|c| c.is_uppercase()) {
            return true;
        }
        false
    };

    for v in &e.variants {
        match &v.kind {
            VariantKind::Plain => {}
            VariantKind::Tuple(fields) => {
                for f in fields {
                    if is_unresolvable(f) {
                        return true;
                    }
                }
            }
            VariantKind::Struct(fields) => {
                for (_, ty) in fields {
                    if is_unresolvable(ty) {
                        return true;
                    }
                }
            }
        }
    }
    false
}

fn write_mod_file(dir: &Path, modules: &[String], include_enums_mod: bool) -> Result<()> {
    let mut out = String::from("// Auto-generated by wrapper_generator. Do not edit.\n\n");
    if include_enums_mod {
        out.push_str("pub mod enums;\npub use enums::*;\n\n");
    }
    for m in modules {
        out.push_str(&format!("mod {m};\npub use {m}::*;\n\n"));
    }
    std::fs::write(dir.join("mod.rs"), &out)?;
    Ok(())
}

fn run_rustfmt(struct_dir: &Path, enums_dir: &Path) {
    let collect_rs = |dir: &Path| -> Vec<PathBuf> {
        std::fs::read_dir(dir)
            .into_iter()
            .flatten()
            .filter_map(|e| e.ok())
            .map(|e| e.path())
            .filter(|p| p.extension().is_some_and(|ext| ext == "rs"))
            .collect()
    };

    let mut files = collect_rs(struct_dir);
    files.extend(collect_rs(enums_dir));

    if files.is_empty() {
        return;
    }

    eprintln!("Running rustfmt on {} files...", files.len());
    let status = std::process::Command::new("rustup")
        .args(["run", "stable", "rustfmt", "--edition=2021"])
        .args(&files)
        .status();

    match status {
        Ok(s) if s.success() => eprintln!("rustfmt: OK"),
        Ok(s) => eprintln!("rustfmt exited with: {}", s),
        Err(e) => eprintln!("rustfmt not available: {}", e),
    }
}

// ── verify ──────────────────────────────────────────────────────────────────

fn cmd_verify(manifest: &Path, overrides_path: &Path) -> Result<()> {
    let AnalysisResult { analyzed, .. } = load_and_analyze(manifest, overrides_path)?;

    println!("=== Wrapper Generator: Verify ===\n");
    println!(
        "Structs: {}, Enums: {}\n",
        analyzed.structs.len(),
        analyzed.enums.len()
    );

    println!("--- Structs ---");
    for s in &analyzed.structs {
        let role = match &s.role {
            StructRole::Standalone => "standalone".to_string(),
            StructRole::Proxy {
                parent_name,
                accessors,
            } => {
                let accs: Vec<&str> = accessors.iter().map(|a| a.parent_method.as_str()).collect();
                format!("proxy({}::{{{}}})", parent_name, accs.join(", "))
            }
        };

        let counts = count_overrides(&s.methods);
        let default_tag = if s.has_default { " [Default]" } else { "" };
        let ctor_tag = if s.constructor.is_some() {
            " [new]"
        } else {
            ""
        };

        println!(
            "  {:<40} {:<35} {:>3} methods (auto={}, skip={}, custom={}, rename={}){}{}",
            s.name,
            role,
            s.methods.len(),
            counts.0,
            counts.1,
            counts.2,
            counts.3,
            default_tag,
            ctor_tag,
        );
    }

    println!("\n--- Enums ---");
    for e in &analyzed.enums {
        let tags = [
            e.has_data_variants().then_some("[data]"),
            e.has_default.then_some("[Default]"),
        ];
        let tags_str: Vec<&str> = tags.into_iter().flatten().collect();
        println!(
            "  {:<40} {:>3} variants  {}",
            e.name,
            e.variants.len(),
            tags_str.join(" ")
        );
    }

    let total_methods: usize = analyzed.structs.iter().map(|s| s.methods.len()).sum();
    let auto_methods: usize = analyzed
        .structs
        .iter()
        .flat_map(|s| s.methods.iter())
        .filter(|m| matches!(m.override_, MethodOverride::Auto))
        .count();

    println!("\n--- Summary ---");
    println!("  Structs:  {}", analyzed.structs.len());
    println!("  Enums:    {}", analyzed.enums.len());
    println!(
        "  Methods:  {} total, {} auto-gen, {} overridden",
        total_methods,
        auto_methods,
        total_methods - auto_methods
    );

    Ok(())
}

/// (auto, skip, custom, rename)
fn count_overrides(methods: &[ir::AnalyzedMethod]) -> (usize, usize, usize, usize) {
    let mut auto = 0;
    let mut skip = 0;
    let mut custom = 0;
    let mut rename = 0;
    for m in methods {
        match &m.override_ {
            MethodOverride::Auto => auto += 1,
            MethodOverride::Skip(_) => skip += 1,
            MethodOverride::Custom(_) => custom += 1,
            MethodOverride::Rename(_) => rename += 1,
        }
    }
    (auto, skip, custom, rename)
}

// ── diff ────────────────────────────────────────────────────────────────────

fn cmd_diff(
    manifest: &Path,
    overrides_path: &Path,
    existing_dir: &Path,
    struct_filter: Option<&str>,
) -> Result<()> {
    let AnalysisResult { analyzed, .. } = load_and_analyze(manifest, overrides_path)?;

    println!("=== Wrapper Generator: Diff ===\n");
    println!(
        "Comparing upstream ({} structs) against: {}\n",
        analyzed.structs.len(),
        existing_dir.display()
    );

    for s in &analyzed.structs {
        if let Some(filter) = struct_filter {
            if s.name != filter {
                continue;
            }
        }

        let snake_name = to_snake_case_filename(&s.name);
        let existing_path = existing_dir.join(format!("{}.rs", snake_name));

        let auto_count = s.auto_methods().count();

        if !existing_path.exists() {
            println!(
                "{}: NO EXISTING FILE — {} auto methods to generate",
                s.name, auto_count
            );
            continue;
        }

        let existing_code = std::fs::read_to_string(&existing_path)
            .with_context(|| format!("Failed to read file: {}", existing_path.display()))?;

        let missing: Vec<_> = s
            .auto_methods()
            .filter(|m| !existing_code.contains(&format!("fn {}", m.name)))
            .collect();

        if missing.is_empty() {
            println!("{}: OK ({} methods all present)", s.name, auto_count);
        } else {
            println!(
                "{}: {} / {} methods missing:",
                s.name,
                missing.len(),
                auto_count
            );
            for m in &missing {
                println!(
                    "  - {} (js: {}, {:?}, {:?})",
                    m.name, m.js_name, m.receiver, m.returns
                );
            }

            // Only output the generated code when --struct_name is specified
            if struct_filter.is_some() {
                println!("\n--- Generated wrapper (full) ---\n");
                let code = codegen::generate_struct_file(s, &codegen::CodegenContext::empty());
                println!("{}", code);
            }
        }
    }

    Ok(())
}
