const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const BUMP_TYPE = process.argv[2] || 'patch'; // patch | minor | major

function bumpSemver(version, type) {
  const [major, minor, patch] = version.split('.').map(Number);
  if (type === 'major') return `${major + 1}.0.0`;
  if (type === 'minor') return `${major}.${minor + 1}.0`;
  return `${major}.${minor}.${patch + 1}`;
}

// 1. Bump package.json
const pkgPath = path.join(__dirname, '..', 'package.json');
const pkg = JSON.parse(fs.readFileSync(pkgPath, 'utf-8'));
const oldVersion = pkg.version;
pkg.version = bumpSemver(oldVersion, BUMP_TYPE);
fs.writeFileSync(pkgPath, JSON.stringify(pkg, null, 2) + '\n');
console.log(`Bumped package.json: ${oldVersion} → ${pkg.version}`);

// 2. (Optional) Write version into the Python script so --version works
const pyPath = path.join(__dirname, '..', 'src', 'aggregate_workbook.py');
let pySrc = fs.readFileSync(pyPath, 'utf-8');
pySrc = pySrc.replace(/__VERSION__ = "[^"]*"/, `__VERSION__ = "${pkg.version}"`);
if (!pySrc.includes('__VERSION__')) {
  pySrc = pySrc.replace('"""', `"""\n__VERSION__ = "${pkg.version}"`);
}
fs.writeFileSync(pyPath, pySrc);

// 3. Rebuild the bundled binary so it matches the new version
console.log('Rebuilding Python binary...');
execSync('npm run build:python', { stdio: 'inherit', cwd: path.join(__dirname, '..') });

// 4. Git commit + tag
execSync('git add -A', { cwd: path.join(__dirname, '..') });
execSync(`git commit -m "chore(release): v${pkg.version}"`, { cwd: path.join(__dirname, '..') });
execSync(`git tag v${pkg.version}`, { cwd: path.join(__dirname, '..') });

console.log(`\n✅ Ready: v${pkg.version}`);
console.log(`   Push with: git push && git push --tags`);