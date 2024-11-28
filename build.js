
import shell from 'shelljs';

if (!shell.which('wasm-pack')) {
	console.error("Run npm install to install all dependencies first");
}

// Clean up any existing build
shell.rm('-rf', 'rust/pkg');

// Create the node output
shell.exec('wasm-pack build ./rust --target nodejs --out-dir pkg/nodejs --release');

// Create the web output
shell.exec('wasm-pack build ./rust --target web --out-dir pkg/web --release');

// Clean up
shell.rm('rust/pkg/*/.gitignore');
