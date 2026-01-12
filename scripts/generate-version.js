const fs = require('fs');
const { execSync } = require('child_process');

// Get git commit hash
let gitCommit = null;
try {
  gitCommit = execSync('git rev-parse HEAD').toString().trim();
  console.log(`Git commit: ${gitCommit.substring(0, 7)}`);
} catch (error) {
  console.log('Git not available or not a git repository');
}

// Get version from manifest.xml
let version = '1.0.0.0';
try {
  const manifest = fs.readFileSync('manifest.xml', 'utf8');
  const match = manifest.match(/<Version>(.*?)<\/Version>/);
  if (match) {
    version = match[1];
    console.log(`Version from manifest: ${version}`);
  }
} catch (error) {
  console.error('Could not read manifest.xml, using default version');
}

// Create helpers directory if it doesn't exist
const helpersDir = 'src/helpers';
if (!fs.existsSync(helpersDir)) {
  fs.mkdirSync(helpersDir, { recursive: true });
}

// Write version file
const versionContent = `// Auto-generated file - do not edit manually
// Generated at: ${new Date().toISOString()}

export const APP_VERSION = "${version}";
export const GIT_COMMIT = ${gitCommit ? `"${gitCommit}"` : 'null'};
export const BUILD_DATE = "${new Date().toISOString()}";
`;

fs.writeFileSync('src/helpers/version.ts', versionContent);
console.log('✓ Version info generated successfully');