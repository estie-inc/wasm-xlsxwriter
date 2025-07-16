use anyhow::{Context, Result};
use std::collections::{HashMap, HashSet};
use std::fs;
use crate::crate_info_extractor::{ExtractedItems, StructInfo};

#[derive(Debug, Clone)]
pub struct StructComparison {
    pub name: String,
    pub status: MigrationStatus,
    pub common_methods: Vec<String>,
    pub wasm_only_methods: Vec<String>,
    pub rust_only_methods: Vec<String>,
    pub common_functions: Vec<String>,
    pub wasm_only_functions: Vec<String>,
    pub rust_only_functions: Vec<String>,
}

#[derive(Debug, Clone)]
pub struct EnumComparison {
    pub name: String,
    pub status: MigrationStatus,
}

#[derive(Debug, PartialEq, Clone)]
pub enum MigrationStatus {
    FullyMigrated,
    PartiallyMigrated,
    NotMigrated,
}

pub fn write_comparison_report(comparison: &ComparisonResults, output_file: &str) -> Result<()> {
    let mut report = String::new();

    report.push_str("# Method Comparison Report: rustxlsxwriter to wasm-xlsxwriter Migration\n\n");

    let mut fully_migrated: Vec<StructComparison> = Vec::new();
    let mut partially_migrated: Vec<StructComparison> = Vec::new();
    let mut not_migrated_structs: Vec<StructComparison> = Vec::new();
    let mut migrated_enums: Vec<EnumComparison> = Vec::new();
    let mut not_migrated_enums: Vec<EnumComparison> = Vec::new();

    let mut migrated_method_count = 0;
    let mut not_migrated_method_count = 0;
    let mut migrated_function_count = 0;
    let mut not_migrated_function_count = 0;
    for struct_comp in &comparison.structs {
        match struct_comp.status {
            MigrationStatus::FullyMigrated => fully_migrated.push(struct_comp.clone()),
            MigrationStatus::PartiallyMigrated => partially_migrated.push(struct_comp.clone()),
            MigrationStatus::NotMigrated => not_migrated_structs.push(struct_comp.clone()),
        }
        migrated_method_count += struct_comp.common_methods.len();
        not_migrated_method_count += struct_comp.rust_only_methods.len();
        migrated_function_count += struct_comp.common_functions.len();
        not_migrated_function_count += struct_comp.rust_only_functions.len();
    }

    for enum_comp in &comparison.enums {
        match enum_comp.status {
            MigrationStatus::FullyMigrated => migrated_enums.push(enum_comp.clone()),
            _ => not_migrated_enums.push(enum_comp.clone()),
        }
    }

    report.push_str("## Summary\n");
    report.push_str(&format!("  - ✅ Fully Migrated Structs: {}\n", fully_migrated.len()));
    report.push_str(&format!("  - ⚠️ Partially Migrated Structs: {}\n", partially_migrated.len()));
    report.push_str(&format!("  - ❌ Not Migrated Structs: {}\n", not_migrated_structs.len()));
    report.push_str(&format!("  - ✅ Migrated Enums: {}\n", migrated_enums.len()));
    report.push_str(&format!("  - ❌ Not Migrated Enums: {}\n", not_migrated_enums.len()));
    report.push_str(&format!("  - ✅ Total Migrated Methods: {}\n", migrated_method_count));
    report.push_str(&format!("  - ❌ Total Not Migrated Methods: {}\n", not_migrated_method_count));
    report.push_str(&format!("  - ✅ Total Migrated Functions: {}\n", migrated_function_count));
    report.push_str(&format!("  - ❌ Total Not Migrated Functions: {}\n", not_migrated_function_count));

    report.push_str("## Details of Structs\n");

    for struct_comp in &fully_migrated {
        report.push_str(&format!("  ### ✅ {}\n", struct_comp.name));
    }

    for struct_comp in partially_migrated.iter().chain(not_migrated_structs.iter()) {
        let status_icon = if struct_comp.status  == MigrationStatus::PartiallyMigrated {
            "⚠️"
        } else {
            "❌"
        };

        report.push_str(&format!("  ### {} {}\n", status_icon, struct_comp.name));

        report.push_str("    Summary\n");
        report.push_str(&format!("      - Migrated methods: {}\n", struct_comp.common_methods.len()));
        report.push_str(&format!("      - Not migrated methods: {}\n", struct_comp.rust_only_methods.len()));
        report.push_str(&format!("      - Migrated functions: {}\n", struct_comp.common_functions.len()));
        report.push_str(&format!("      - Not migrated functions: {}\n", struct_comp.rust_only_functions.len()));

        if !struct_comp.rust_only_methods.is_empty() {
            report.push_str("    ❌ Methods Not Yet Migrated\n");
            for method in &struct_comp.rust_only_methods {
                report.push_str(&format!("      - {}\n", method));
            }
        }

        if !struct_comp.rust_only_functions.is_empty() {
            report.push_str("    ❌ Functions Not Yet Migrated\n");
            for function in &struct_comp.rust_only_functions {
                report.push_str(&format!("      - {}\n", function));
            }
        }
    }

    report.push_str("\n## ✅ Migrated Enums\n");
    for enum_comp in &migrated_enums {
        report.push_str(&format!("  - {}\n", enum_comp.name));
    }

    report.push_str("\n## ❌ Not Migrated Enums\n");
    for enum_comp in &not_migrated_enums {
        report.push_str(&format!("  - {}\n", enum_comp.name));
    }

    fs::write(output_file, report)
        .context(format!("Failed to write comparison report to {}", output_file))?;

    Ok(())
}

pub struct ComparisonResults {
    pub structs: Vec<StructComparison>,
    pub enums: Vec<EnumComparison>,
}

pub fn compare_methods(
    wasm_items: &ExtractedItems,
    rust_items: &ExtractedItems,
) -> ComparisonResults {
    let mut struct_comparisons = Vec::new();
    let mut enum_comparisons = Vec::new();

    let wasm_structs_map: HashMap<String, &StructInfo> = wasm_items.structs.iter()
        .map(|s| (s.name.clone(), s))
        .collect();
    let rust_structs_map: HashMap<String, &StructInfo> = rust_items.structs.iter()
        .map(|s| (s.name.clone(), s))
        .collect();

    let wasm_enums: HashSet<String> = wasm_items.enums.iter().map(|e| e.name.clone()).collect();
    let rust_enums: HashSet<String> = rust_items.enums.iter().map(|e| e.name.clone()).collect();

    let all_enums: HashSet<String> = wasm_enums.union(&rust_enums).cloned().collect();
    for enum_name in all_enums {
        let in_wasm_enum = wasm_enums.contains(&enum_name) || wasm_structs_map.contains_key(&enum_name);
        let in_rust_enum = rust_enums.contains(&enum_name);

        let status = if in_wasm_enum && in_rust_enum {
            MigrationStatus::FullyMigrated
        } else {
            MigrationStatus::NotMigrated
        };

        enum_comparisons.push(EnumComparison {
            name: enum_name,
            status,
        });
    }

    let all_structs: HashSet<String> = wasm_structs_map.keys().chain(rust_structs_map.keys())
        .cloned()
        .collect();
    for struct_name in all_structs {
        let wasm_struct = wasm_structs_map.get(&struct_name);
        let rust_struct = rust_structs_map.get(&struct_name);

        let wasm_struct_methods: HashSet<String> = wasm_struct
            .map(|s| s.methods.iter().map(|m| m.method_name.clone()).collect())
            .unwrap_or_default();

        let rust_struct_methods: HashSet<String> = rust_struct
            .map(|s| s.methods.iter().map(|m| m.method_name.clone()).collect())
            .unwrap_or_default();

        let wasm_struct_functions: HashSet<String> = wasm_struct
            .map(|s| s.functions.iter().map(|f| f.name.clone()).collect())
            .unwrap_or_default();

        let rust_struct_functions: HashSet<String> = rust_struct
            .map(|s| s.functions.iter().map(|f| f.name.clone()).collect())
            .unwrap_or_default();

        if rust_struct_methods.is_empty() && rust_struct_functions.is_empty() {
            continue;
        }

        let mut common_methods: Vec<String> = wasm_struct_methods
            .intersection(&rust_struct_methods)
            .cloned()
            .collect();
        common_methods.sort();

        let mut wasm_only_methods: Vec<String> = wasm_struct_methods
            .difference(&rust_struct_methods)
            .cloned()
            .collect();
        wasm_only_methods.sort();

        let mut rust_only_methods: Vec<String> = rust_struct_methods
            .difference(&wasm_struct_methods)
            .cloned()
            .collect();
        rust_only_methods.sort();

        let mut common_functions: Vec<String> = wasm_struct_functions
            .intersection(&rust_struct_functions)
            .cloned()
            .collect();
        common_functions.sort();

        let mut wasm_only_functions: Vec<String> = wasm_struct_functions
            .difference(&rust_struct_functions)
            .cloned()
            .collect();
        wasm_only_functions.sort();

        let mut rust_only_functions: Vec<String> = rust_struct_functions
            .difference(&wasm_struct_functions)
            .cloned()
            .collect();
        rust_only_functions.sort();

        let status = if common_methods.is_empty() && common_functions.is_empty() {
            MigrationStatus::NotMigrated
        } else if rust_only_methods.is_empty() && rust_only_functions.is_empty() {
            MigrationStatus::FullyMigrated
        } else {
            MigrationStatus::PartiallyMigrated
        };

        struct_comparisons.push(StructComparison {
            name: struct_name,
            status,
            common_methods,
            wasm_only_methods,
            rust_only_methods,
            common_functions,
            wasm_only_functions,
            rust_only_functions,
        });
    }

    struct_comparisons.sort_by(|a, b| a.name.cmp(&b.name));
    enum_comparisons.sort_by(|a, b| a.name.cmp(&b.name));

    ComparisonResults {
        structs: struct_comparisons,
        enums: enum_comparisons,
    }
}
