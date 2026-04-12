const fs = require('fs');
const path = require('path');

const srcDir = path.resolve(__dirname, '..', 'src', 'renderer');
const destDir = path.resolve(__dirname, '..', 'dist', 'renderer');

function copyRecursive(source, destination) {
  if (!fs.existsSync(destination)) {
    fs.mkdirSync(destination, { recursive: true });
  }

  const entries = fs.readdirSync(source, { withFileTypes: true });
  for (const entry of entries) {
    const sourcePath = path.join(source, entry.name);
    const destinationPath = path.join(destination, entry.name);

    if (entry.isDirectory()) {
      copyRecursive(sourcePath, destinationPath);
      continue;
    }

    fs.copyFileSync(sourcePath, destinationPath);
  }
}

if (!fs.existsSync(srcDir)) {
  throw new Error(`Renderer source folder not found: ${srcDir}`);
}

copyRecursive(srcDir, destDir);
console.log('Renderer assets copied to dist/renderer');
