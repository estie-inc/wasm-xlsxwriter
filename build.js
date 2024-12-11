
import shell from 'shelljs';

if (!shell.which('wasm-pack')) {
	console.error("Run npm install to install all dependencies first");
}

// Clean up any existing build
shell.rm('-rf', 'pkg.*');

// Create the node output
shell.exec('wasm-pack build ./rust --target nodejs --out-dir ../pkg.node --release');

// Create the web output
shell.exec('wasm-pack build ./rust --target web --out-dir ../pkg.web --release');

// Create the bundler output
// shell.exec('wasm-pack build ./rust --target bundler --out-dir ../pkg.bundler --release');

// Clean up
shell.rm('pkg.*/.gitignore');
