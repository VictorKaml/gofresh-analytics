#!/usr/bin/env node
/**
 * build-python.js
 *
 * Compiles src/aggregate_workbook.py into a standalone, self-contained
 * executable folder (via PyInstaller --onedir) at build-resources/python/.
 * That directory is picked up by electron-builder's `extraResources`
 * config and shipped inside the packaged app, so end users never need
 * Python installed.
 *
 * --onedir (not --onefile) is deliberate: --onefile re-extracts the whole
 * bundled Python runtime to a temp directory on *every single launch*,
 * which is fine for a one-shot CLI tool but is a real cost here since this
 * binary runs on every workbook load (startup, "Open File", drag-and-drop
 * reload...). --onedir ships the already-extracted runtime as a folder, so
 * repeat launches start in well under a second instead of several.
 *
 * IMPORTANT: PyInstaller does not cross-compile. Run this script on the
 * same OS you're packaging for (Windows -> run on Windows, etc.) — the
 * npm run dist target is Windows, so this normally needs to run on a
 * Windows machine (or CI runner) before `npm run dist`.
 *
 * Requirements (build machine only):
 *   pip install -r requirements.txt      (openpyxl + pyinstaller)
 */

const { execFileSync } = require('child_process');
const path = require('path');
const fs = require('fs');
const os = require('os');

const ROOT = path.join(__dirname, '..');
const SRC_SCRIPT = path.join(ROOT, 'src', 'aggregate_workbook.py');
const DIST_PATH = path.join(ROOT, 'build-resources', 'python');
const WORK_PATH = path.join(os.tmpdir(), 'gofresh-pyi-work');
const SPEC_PATH = path.join(os.tmpdir(), 'gofresh-pyi-spec');

const PYTHON_CANDIDATES = process.platform === 'win32'
  ? ['python', 'py']
  : ['python3', 'python'];

function resolvePython() {
  for (const candidate of PYTHON_CANDIDATES) {
    try {
      execFileSync(candidate, ['--version'], { stdio: 'ignore' });
      return candidate;
    } catch (_) {
      // try next
    }
  }
  console.error(
    `Could not find a Python interpreter (tried: ${PYTHON_CANDIDATES.join(', ')}).\n` +
    `Install Python 3, then run: pip install -r requirements.txt`
  );
  process.exit(1);
}

function main() {
  const python = resolvePython();

  console.log(`Using interpreter: ${python}`);
  console.log('Building standalone aggregator binary with PyInstaller...');

  fs.rmSync(DIST_PATH, { recursive: true, force: true });

  execFileSync(python, [
    '-m', 'PyInstaller',
    '--onedir',
    '--name', 'aggregate_workbook',
    '--distpath', DIST_PATH,
    '--workpath', WORK_PATH,
    '--specpath', SPEC_PATH,
    // python_calamine wraps a compiled Rust extension (_python_calamine.*
    // .pyd), not pure Python — plain --collect-submodules only walks .py
    // files, so it can miss the binary. --collect-all grabs submodules,
    // binaries, and data together, which is what a compiled extension needs.
    '--collect-all', 'python_calamine',
    // On a build machine with a data-science stack installed globally,
    // PyInstaller's static analyzer can pull in unrelated heavy packages
    // "just in case" if anything on the import graph even optionally
    // references them, ballooning the bundle for code that's never
    // executed. aggregate_workbook.py needs none of these — exclude them.
    '--exclude-module', 'pandas',
    '--exclude-module', 'numpy',
    '--exclude-module', 'scipy',
    '--exclude-module', 'matplotlib',
    '--exclude-module', 'PIL',
    '--exclude-module', 'gi',
    '--exclude-module', 'psutil',
    '--exclude-module', 'lxml',
    '--exclude-module', 'tkinter',
    '--exclude-module', 'IPython',
    '--exclude-module', 'pytest',
    '--clean',
    '--noconfirm',
    SRC_SCRIPT,
  ], { stdio: 'inherit', cwd: ROOT });

  const exeName = process.platform === 'win32' ? 'aggregate_workbook.exe' : 'aggregate_workbook';
  // --onedir produces build-resources/python/aggregate_workbook/aggregate_workbook(.exe)
  // — a folder containing the executable plus its bundled runtime/libs.
  const exePath = path.join(DIST_PATH, 'aggregate_workbook', exeName);
  if (!fs.existsSync(exePath)) {
    console.error(`Build finished but expected output not found at ${exePath}`);
    process.exit(1);
  }

  console.log(`\nBuilt: ${exePath}`);
  console.log('This is picked up automatically by "npm start" (dev) and');
  console.log('bundled into the packaged app by "npm run dist" (extraResources).');
}

main();