use anyhow::{Context, Result};
use std::fs;
use std::path::{Path, PathBuf};
use std::process;

pub fn download_repository(repo_url: &str) -> Result<PathBuf> {
    let repos_dir = ensure_repos_directory()?;

    let name = repo_url
        .split('/')
        .last()
        .and_then(|s| s.strip_suffix(".git"))
        .context("Could not find repository name")?;

    let repo_path = repos_dir.join(name);
    if repo_path.exists() {
        println!("Using existing repository: {}", name);
        return Ok(repo_path);
    }
    clone_repository(repo_url, &repos_dir)?;

    Ok(repo_path)
}

fn ensure_repos_directory() -> Result<PathBuf> {
    let repos_dir = Path::new("./target/repos");
    fs::create_dir_all(repos_dir).context("Failed to create repositories directory")?;
    Ok(repos_dir.to_path_buf())
}

fn clone_repository(repo_url: &str, repos_dir: &Path) -> Result<()> {
    println!("Downloading repository: {}", repo_url);

    let status = process::Command::new("git")
        .args(["clone", repo_url])
        .current_dir(repos_dir)
        .status()
        .context("Failed to execute git clone command")?;

    if !status.success() {
        anyhow::bail!("Git clone command failed with non-zero exit status");
    }

    Ok(())
}
