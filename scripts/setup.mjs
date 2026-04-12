import { execSync } from "node:child_process";
import { existsSync, mkdirSync, symlinkSync, readFileSync, unlinkSync, writeFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = join(dirname(fileURLToPath(import.meta.url)), "..");
const CACHE_DIR = join(ROOT, "node_modules", ".cache");
const BIN_DIR = join(ROOT, "node_modules", ".bin");
const BINARYEN_VERSION = 123;

function getPlatform() {
  const key = `${process.platform}-${process.arch}`;
  const wasmBindgen = {
    "darwin-arm64": "aarch64-apple-darwin",
    "darwin-x64": "x86_64-apple-darwin",
    "linux-x64": "x86_64-unknown-linux-musl",
    "linux-arm64": "aarch64-unknown-linux-musl",
  }[key];
  const binaryen = {
    "darwin-arm64": "arm64-macos",
    "darwin-x64": "x86_64-macos",
    "linux-x64": "x86_64-linux",
    "linux-arm64": "aarch64-linux",
  }[key];
  if (!wasmBindgen) throw new Error(`Unsupported platform: ${key}`);
  return { wasmBindgen, binaryen };
}

function getWasmBindgenVersion() {
  const lockfile = readFileSync(join(ROOT, "rust", "Cargo.lock"), "utf-8");
  const match = lockfile.match(/name = "wasm-bindgen"\nversion = "([^"]+)"/);
  if (!match) throw new Error("wasm-bindgen version not found in rust/Cargo.lock");
  return match[1];
}

async function downloadAndExtract(url, destDir) {
  const res = await fetch(url);
  if (!res.ok) throw new Error(`Download failed: ${url} (${res.status})`);
  const buffer = Buffer.from(await res.arrayBuffer());
  mkdirSync(destDir, { recursive: true });
  const tmp = join(destDir, "_download.tar.gz");
  writeFileSync(tmp, buffer);
  execSync(`tar xzf "${tmp}" -C "${destDir}"`);
  unlinkSync(tmp);
}

function link(target, name) {
  const linkPath = join(BIN_DIR, name);
  try { unlinkSync(linkPath); } catch {}
  mkdirSync(BIN_DIR, { recursive: true });
  symlinkSync(target, linkPath);
}

async function setupWasmBindgen() {
  const version = getWasmBindgenVersion();
  const platform = getPlatform().wasmBindgen;
  const dirName = `wasm-bindgen-${version}-${platform}`;
  const binPath = join(CACHE_DIR, dirName, "wasm-bindgen");

  if (!existsSync(binPath)) {
    const url = `https://github.com/wasm-bindgen/wasm-bindgen/releases/download/${version}/wasm-bindgen-${version}-${platform}.tar.gz`;
    console.log(`Downloading wasm-bindgen ${version}...`);
    await downloadAndExtract(url, CACHE_DIR);
  }

  link(binPath, "wasm-bindgen");
  console.log(`wasm-bindgen ${version} ok`);
}

async function setupWasmOpt() {
  const platform = getPlatform().binaryen;
  const dirName = `binaryen-version_${BINARYEN_VERSION}`;
  const binPath = join(CACHE_DIR, dirName, "bin", "wasm-opt");

  if (!existsSync(binPath)) {
    const url = `https://github.com/WebAssembly/binaryen/releases/download/version_${BINARYEN_VERSION}/binaryen-version_${BINARYEN_VERSION}-${platform}.tar.gz`;
    console.log(`Downloading binaryen ${BINARYEN_VERSION}...`);
    await downloadAndExtract(url, CACHE_DIR);
  }

  link(binPath, "wasm-opt");
  console.log(`wasm-opt ${BINARYEN_VERSION} ok`);
}

await setupWasmBindgen();
await setupWasmOpt();
