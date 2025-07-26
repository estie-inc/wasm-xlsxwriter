mod repo_downloader;
mod report_writer;

use crate::repo_downloader::download_repository;
use crate::report_writer::write_comparison_report;
use anyhow::{Context, Result};
use clap::Parser;
use crate_inspector::{Crate, CrateBuilder};

#[derive(Parser, Debug)]
#[command(author, version, about, long_about = None)]
struct Args {
    /// Output file path for the comparison report
    #[arg(short, long, default_value = "method_comparison.md")]
    output: String,
}

/// A tool to compare methods implemented in wasm-xlsxwriter and rustxlsxwriter

fn main() -> Result<()> {
    let args = Args::parse();

    let wasm_manifest_path = "../../Cargo.toml";
    let wasm_crate = get_crate_info(wasm_manifest_path, true)?;

    let repo_path = download_repository("https://github.com/jmcnamara/rust_xlsxwriter.git")?;
    let rust_crate = get_crate_info(repo_path.join("Cargo.toml").to_str().unwrap(), false)?;

    let comparison = report_writer::compare_methods(&wasm_crate, &rust_crate);
    write_comparison_report(&comparison, &args.output)?;

    println!("Comparison report written to {}", args.output);
    Ok(())
}

fn get_crate_info(manifest_path: &str, document_private: bool) -> Result<Crate> {
    Ok(CrateBuilder::default()
        .manifest_path(manifest_path)
        .document_private_items(document_private)
        .build()
        .context("Failed to build crate inspector")?)
}
