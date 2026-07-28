# Repository Guidelines

## Project Structure & Module Organization

This `clasp`-managed Google Apps Script project tracks submitted items from barcode scans. Deployable files live at the root because `.clasp.json` uses an empty `rootDir`:

- `コード.js`: spreadsheet menu setup, scan-list initialization, barcode generation, and submission checks.
- `appsscript.json`: V8 runtime, Japan timezone, logging, and spreadsheet macro configuration.
- `.clasp.json`: local Apps Script project connection; `clasp status` excludes it from deployment.
- `tests/`: Node.js unit tests for pure barcode and matching logic.
- `.github/workflows/ci.yml`: syntax and unit-test checks for pushes and pull requests.

Keep deployable Apps Script modules at the root unless the clasp layout is deliberately changed. Update `.claspignore` whenever adding development-only JavaScript or JSON files.

## Build, Test, and Development Commands

No compilation or package manager is configured. Use:

```powershell
node --check ".\コード.js"   # Validate JavaScript syntax locally
node --test tests/*.test.js  # Run the unit-test suite
clasp status                 # Confirm exactly which files will be uploaded
clasp pull                   # Refresh local files from Apps Script
clasp push                   # Upload tracked source and manifest files
clasp open                   # Open the bound Apps Script project
```

Review local changes before `clasp pull`, which can replace local content. Treat `clasp push` as a deployment action and run it only after validation.

## Coding Style & Naming Conventions

Use two-space indentation, braces for all control blocks, semicolons, and single-quoted strings. Prefer `const` and `let` for new code; retain existing public function names such as `onOpen`, `syokika`, and `barcodelabelcheck` because spreadsheet menus and macros reference them. Use descriptive camelCase names for new functions and variables. Batch spreadsheet reads and writes where practical instead of calling `getRange()` repeatedly inside loops.

## Testing Guidelines

Use the built-in `node:test` framework and name test files `*.test.js`. Cover normalization, empty input, duplicate and unknown IDs, date-column reuse, and batch-output helpers. Run syntax and unit tests before every push. Also test Spreadsheet service integration in a duplicate spreadsheet containing the expected `読み込み` and `バーコード作成` sheets. Do not test with production student data.

## Commit & Pull Request Guidelines

Create work from a GitHub Issue and include its number in the branch name, for example `issue-12-fix-unknown-barcodes`. Use focused imperative commits such as `Improve submission scan safety`. Pull requests must include `Closes #<issue>`, describe affected spreadsheet workflows, list automated and manual checks, and include screenshots for visible sheet or menu changes.

## Security & Configuration

Do not publish `.clasp.json`, credentials, spreadsheet contents, or personally identifiable submission data. Review external barcode-service URLs before changing parameters or providers.
