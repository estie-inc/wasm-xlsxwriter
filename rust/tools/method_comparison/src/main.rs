mod crate_info_extractor;
mod repo_downloader;
mod report_writer;

use anyhow::Result;
use clap::Parser;
use crate::crate_info_extractor::{get_crate_info, extract_crate_items};
use crate::repo_downloader::download_repository;
use crate::report_writer::write_comparison_report;

#[derive(Parser, Debug)]
#[command(author, version, about, long_about = None)]
struct Args {
    /// Output file path for the comparison report
    #[arg(short, long, default_value = "method_comparison.md")]
    output: String,
}

/// A tool to compare methods implemented in wasm-xlsxwriter and rustxlsxwriter

fn main() -> Result<()> {
    // Parse command line arguments
    let args = Args::parse();

    // Hardcoded values
    let wasm_manifest_path = "../../Cargo.toml";
    let wasm_crate = get_crate_info(wasm_manifest_path)?;

    let repo_path = download_repository("https://github.com/jmcnamara/rust_xlsxwriter.git")?;
    let rust_crate = get_crate_info(repo_path.join("Cargo.toml").to_str().unwrap())?;

    let wasm_items = extract_crate_items(&wasm_crate);
    let rust_items = extract_crate_items(&rust_crate);

    // println!("wasm_items: {:#?}", wasm_items);
    // println!("rust_items: {:#?}", rust_items);

    let comparison = report_writer::compare_methods(&wasm_items, &rust_items);

    write_comparison_report(&comparison, &args.output)?;

    println!("Comparison report written to {}", args.output);
    Ok(())
}
