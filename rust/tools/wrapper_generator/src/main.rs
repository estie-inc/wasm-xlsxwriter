use anyhow::{Context, Result};
use clap::Parser;
use crate_inspector::{CrateBuilder, StructItem, EnumItem};
use ruast::*;
use std::path::PathBuf;

mod common;
mod enum_wrapper;
mod method_generator;
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

    /// Name of the struct to generate wrappers for
    #[arg(short, long)]
    struct_name: Option<String>,

    /// Name of the enum to generate wrappers for
    #[arg(short, long)]
    enum_name: Option<String>,

    /// Output file path (optional, defaults to stdout)
    #[arg(short, long)]
    output: Option<PathBuf>,
}

fn main() -> Result<()> {
    let args = Args::parse();

    if args.struct_name.is_none() && args.enum_name.is_none() {
        return Err(anyhow::anyhow!("Either --struct-name or --enum-name must be provided"));
    }

    if args.struct_name.is_some() && args.enum_name.is_some() {
        return Err(anyhow::anyhow!("Only one of --struct-name or --enum-name can be provided"));
    }

    let krate = get_crate_info(&args.manifest_path, false).context("Failed to get crate info")?;
    let mut output_krate = Crate::new();

    if let Some(struct_name) = args.struct_name {
        let struct_info = krate
            .all_structs()
            .find(|s| s.name() == struct_name)
            .context(format!("Struct '{}' not found", struct_name))?;

        generate_struct_wrapper_output(&mut output_krate, &struct_info);
    } else if let Some(enum_name) = args.enum_name {
        let enum_info = krate
            .all_enums()
            .find(|e| e.name() == enum_name)
            .context(format!("Enum '{}' not found", enum_name))?;

        generate_enum_wrapper_output(&mut output_krate, &enum_info);
    }

    println!("{}", output_krate);

    Ok(())
}


fn get_crate_info(manifest_path: &str, document_private: bool) -> Result<crate_inspector::Crate> {
    Ok(CrateBuilder::default()
        .manifest_path(manifest_path)
        .document_private_items(document_private)
        .build()
        .context("Failed to build crate inspector")?)
}
