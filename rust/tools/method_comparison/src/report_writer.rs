use anyhow::{Context, Result};
use std::collections::HashSet;
use std::fs;
use crate::crate_info_extractor::ExtractedItems;

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

#[derive(Debug, PartialEq, Clone)]
pub enum MigrationStatus {
    FullyMigrated,
    PartiallyMigrated,
    NotMigrated,
}

pub fn write_comparison_report(comparison: &Vec<StructComparison>, output_file: &str) -> Result<()> {
    let mut report = String::new();

    // Method Comparison Report: rustxlsxwriter to wasm-xlsxwriter Migration
    report.push_str("# Method Comparison Report: rustxlsxwriter to wasm-xlsxwriter Migration\n\n");

    // Sort structs by name for consistent output
    let mut sorted_structs = comparison.clone();
    sorted_structs.sort_by(|a, b| a.name.cmp(&b.name));

    // Group structs by migration status
    let mut fully_migrated = Vec::new();
    let mut partially_migrated = Vec::new();
    let mut not_migrated = Vec::new();

    for struct_comp in &sorted_structs {
        match struct_comp.status {
            MigrationStatus::FullyMigrated => fully_migrated.push(struct_comp.clone()),
            MigrationStatus::PartiallyMigrated => partially_migrated.push(struct_comp.clone()),
            MigrationStatus::NotMigrated => not_migrated.push(struct_comp.clone()),
        }
    }

    // Migrated Structs
    report.push_str("## ✅ Migrated Structs\n");
    if !fully_migrated.is_empty() {
        for struct_comp in &fully_migrated {
            report.push_str(&format!("  - {}\n", struct_comp.name));
        }
    } else {
        report.push_str("  - None\n");
    }
    report.push_str("\n");

    // Partially Migrated Structs
    report.push_str("## ⚠️ Partially Migrated Structs\n");
    if !partially_migrated.is_empty() {
        for struct_comp in &partially_migrated {
            report.push_str(&format!("  ### {}\n", struct_comp.name));

            // Summary
            report.push_str("    #### Summary\n");
            report.push_str(&format!("      - Migrated methods: {}\n", struct_comp.common_methods.len()));
            report.push_str(&format!("      - Not migrated methods: {}\n", struct_comp.rust_only_methods.len()));
            report.push_str(&format!("      - Migrated functions: {}\n", struct_comp.common_functions.len()));
            report.push_str(&format!("      - Not migrated functions: {}\n", struct_comp.rust_only_functions.len()));

            // Methods Not Yet Migrated
            if !struct_comp.rust_only_methods.is_empty() {
                report.push_str("    #### ❌ Methods Not Yet Migrated\n");
                for method in &struct_comp.rust_only_methods {
                    report.push_str(&format!("      - {}\n", method));
                }
            }

            // Functions Not Yet Migrated
            if !struct_comp.rust_only_functions.is_empty() {
                report.push_str("    #### ❌ Functions Not Yet Migrated\n");
                for function in &struct_comp.rust_only_functions {
                    report.push_str(&format!("      - {}\n", function));
                }
            }
        }
    } else {
        report.push_str("  - None\n");
    }
    report.push_str("\n");

    // Not Migrated Structs
    report.push_str("## ❌ Not Migrated Structs\n");
    if not_migrated.is_empty() {
        report.push_str("  - None\n");
    } else {
        for struct_comp in &not_migrated {
            report.push_str(&format!("  ### {}\n", struct_comp.name));

            // Summary
            report.push_str("    #### Summary\n");
            report.push_str(&format!("      - Migrated methods: {}\n", struct_comp.common_methods.len()));
            report.push_str(&format!("      - Not migrated methods: {}\n", struct_comp.rust_only_methods.len()));
            report.push_str(&format!("      - Migrated functions: {}\n", struct_comp.common_functions.len()));
            report.push_str(&format!("      - Not migrated functions: {}\n", struct_comp.rust_only_functions.len()));

            // Methods Not Yet Migrated
            if !struct_comp.rust_only_methods.is_empty() {
                report.push_str("    #### ❌ Methods Not Yet Migrated\n");
                for method in &struct_comp.rust_only_methods {
                    report.push_str(&format!("      - {}\n", method));
                }
            }

            // Functions Not Yet Migrated
            if !struct_comp.rust_only_functions.is_empty() {
                report.push_str("    #### ❌ Functions Not Yet Migrated\n");
                for function in &struct_comp.rust_only_functions {
                    report.push_str(&format!("      - {}\n", function));
                }
            }
        }
    }

    fs::write(output_file, report)
        .context(format!("Failed to write comparison report to {}", output_file))?;

    Ok(())
}

pub fn compare_methods(
    wasm_items: &ExtractedItems,
    rust_items: &ExtractedItems,
) -> Vec<StructComparison> {
    let wasm_structs: HashSet<String> = wasm_items.structs.iter().map(|s| s.name.clone()).collect();
    let rust_structs: HashSet<String> = rust_items.structs.iter().map(|s| s.name.clone()).collect();

    // Get the union of all struct names
    let all_structs: HashSet<String> = wasm_structs.union(&rust_structs).cloned().collect();

    // Convert to sorted Vec for consistent iteration order
    let mut all_structs_vec: Vec<String> = all_structs.into_iter().collect();
    all_structs_vec.sort();

    let mut struct_comparisons = Vec::new();

    // Process all structs in a single loop
    for struct_name in all_structs_vec {
        let in_wasm = wasm_structs.contains(&struct_name);
        let _in_rust = rust_structs.contains(&struct_name);

        // Get methods for this struct from both WASM and Rust
        let wasm_struct_methods: HashSet<String> = wasm_items.structs.iter()
            .find(|s| s.name == struct_name)
            .map(|s| s.methods.iter().map(|m| m.method_name.clone()).collect())
            .unwrap_or_default();

        let rust_struct_methods: HashSet<String> = rust_items.structs.iter()
            .find(|s| s.name == struct_name)
            .map(|s| s.methods.iter().map(|m| m.method_name.clone()).collect())
            .unwrap_or_default();

        // Get functions for this struct from both WASM and Rust
        let wasm_struct_functions: HashSet<String> = wasm_items.structs.iter()
            .find(|s| s.name == struct_name)
            .map(|s| s.functions.iter().map(|f| f.name.clone()).collect())
            .unwrap_or_default();

        let rust_struct_functions: HashSet<String> = rust_items.structs.iter()
            .find(|s| s.name == struct_name)
            .map(|s| s.functions.iter().map(|f| f.name.clone()).collect())
            .unwrap_or_default();

        // Calculate intersections and differences
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

        // Determine migration status based on presence in WASM/Rust and method/function comparison
        // We are migrating from Rust to WASM
        let status = if !in_wasm {
            // Rust-only struct - not migrated at all
            MigrationStatus::NotMigrated
        } else if wasm_only_methods.is_empty() && wasm_only_functions.is_empty() && !common_methods.is_empty() {
            // All methods/functions in WASM are also in Rust - fully migrated
            MigrationStatus::FullyMigrated
        } else if !common_methods.is_empty() || !common_functions.is_empty() {
            // Some methods/functions are common - partially migrated
            MigrationStatus::PartiallyMigrated
        } else {
            MigrationStatus::NotMigrated
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

    // Sort struct_comparisons by name for consistent results
    struct_comparisons.sort_by(|a, b| a.name.cmp(&b.name));
    struct_comparisons
}
