use anyhow::{Context, Result};
use clap::Parser;
use crate_inspector::{CrateBuilder, CrateItem};
use ruast::*;
use std::path::PathBuf;

mod common;
mod enum_wrapper;
mod method_wrapper;
mod struct_wrapper;
mod utils;

use crate::enum_wrapper::generate_enum_wrapper_output;
use crate::struct_wrapper::generate_struct_wrapper_output;

/// Tool to automatically generate wasm_xlsxwriter wrapper methods from rust_xlsxwriter
#[derive(Parser, Debug)]
#[command(author, version, about, long_about = None)]
struct Args {
    /// Path to the Cargo.toml of the crate to extract methods from
    #[arg(short, long)]
    manifest_path: String,

    /// Output directory path
    #[arg(short, long)]
    output_dir: PathBuf,
}

fn main() -> Result<()> {
    let args = Args::parse();

    let krate = get_crate_info(&args.manifest_path, false).context("Failed to get crate info")?;

    // Create output directory if it doesn't exist
    std::fs::create_dir_all(&args.output_dir)
        .context(format!("Failed to create output directory: {:?}", args.output_dir))?;

    for struct_info in krate.all_structs() {
        let struct_krate = generate_struct_wrapper_output(&struct_info);
        
        let module_path = struct_info.module().map(|m| m.name().to_string()).unwrap_or_default();
        let output_path = get_output_path(&args.output_dir, &module_path, &struct_info.name())?;
        write_generated_file(&output_path, &struct_krate.to_string())?;
        
        println!("Generated struct wrapper: {:?} (module: {})", output_path, module_path);
    }

    for enum_info in krate.all_enums() {
        let enum_krate = generate_enum_wrapper_output(&enum_info);
        
        let module_path = enum_info.module().map(|m| m.name().to_string()).unwrap_or_default();
        let output_path = get_output_path(&args.output_dir, &module_path, &enum_info.name())?;
        write_generated_file(&output_path, &enum_krate.to_string())?;
        
        println!("Generated enum wrapper: {:?} (module: {})", output_path, module_path);
    }

    println!("All wrappers generated in directory: {:?}", args.output_dir);

    Ok(())
}


fn get_crate_info(manifest_path: &str, document_private: bool) -> Result<crate_inspector::Crate> {
    Ok(CrateBuilder::default()
        .manifest_path(manifest_path)
        .document_private_items(document_private)
        .build()
        .context("Failed to build crate inspector")?)
}

fn get_output_path(base_dir: &PathBuf, module_path: &str, item_name: &str) -> Result<PathBuf> {
    // For items in a module, create subdirectory structure
    // e.g., format.rs -> FormatAlign should go to format/format_align.rs
    let module_dir = if module_path.is_empty() {
        base_dir.clone()
    } else {
        // Create subdirectory named after the module
        base_dir.join(module_path)
    };
    
    // Convert item name like "FormatAlign" to snake_case filename
    let filename = format!("{}.rs", to_snake_case(item_name));
    Ok(module_dir.join(filename))
}

fn to_snake_case(s: &str) -> String {
    let mut result = String::new();
    for (i, c) in s.chars().enumerate() {
        if i > 0 && c.is_uppercase() {
            result.push('_');
        }
        result.push(c.to_lowercase().next().unwrap());
    }
    result
}

fn write_generated_file(path: &PathBuf, content: &str) -> Result<()> {
    // Create parent directories if they don't exist
    if let Some(parent) = path.parent() {
        std::fs::create_dir_all(parent)
            .context(format!("Failed to create directory: {:?}", parent))?;
    }
    
    std::fs::write(path, content)
        .context(format!("Failed to write file: {:?}", path))?;
    
    Ok(())
}
