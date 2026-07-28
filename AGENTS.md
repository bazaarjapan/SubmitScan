# Repository Guidelines

## Project Structure & Module Organization

This is a small, `clasp`-managed Google Apps Script project for tracking submitted items from barcode scans. Deployable files live at the repository root because `.clasp.json` uses an empty `rootDir`:

- `コード.js`: spreadsheet menu setup, scan-list initialization, barcode generation, and submission checks.
- `appsscript.json`: V8 runtime, Japan timezone, logging, and spreadsheet macro configuration.
- `.clasp.json`: local Apps Script project connection; `clasp status` excludes it from deployment.

There are currently no separate test, asset, or build directories. Keep new Apps Script modules at the root unless the clasp layout is deliberately changed.

## Build, Test, and Development Commands

No compilation step or package manager is configured. Use:

```powershell
node --check ".\コード.js"   # Validate JavaScript syntax locally
clasp status                 # Confirm exactly which files will be uploaded
clasp pull                   # Refresh local files from Apps Script
clasp push                   # Upload tracked source and manifest files
clasp open                   # Open the bound Apps Script project
```

Review local changes before `clasp pull`, which can replace local content. Treat `clasp push` as a deployment action and run it only after validation.

## Coding Style & Naming Conventions

Use two-space indentation, braces for all control blocks, semicolons, and single-quoted strings. Prefer `const` and `let` for new code; retain existing public function names such as `onOpen`, `syokika`, and `barcodelabelcheck` because spreadsheet menus and macros reference them. Use descriptive camelCase names for new functions and variables. Batch spreadsheet reads and writes where practical instead of calling `getRange()` repeatedly inside loops.

## Testing Guidelines

No automated test framework or coverage threshold exists. Run `node --check` and test in a duplicate spreadsheet containing the expected `読み込み` and `バーコード作成` sheets. Verify menu creation, initialization, barcode formulas, duplicate or unknown scans, empty ranges, and the dated result column. Do not test with production student data.

## Commit & Pull Request Guidelines

This directory has no local Git metadata, so no repository-specific commit convention can be inferred. If maintained in Git, use focused imperative commits such as `Fix unknown barcode handling`. Pull requests should describe affected spreadsheet workflows, list manual checks, link an issue when available, and include screenshots for visible sheet or menu changes.

## Security & Configuration

Do not publish `.clasp.json`, credentials, spreadsheet contents, or personally identifiable submission data. Review external barcode-service URLs before changing parameters or providers.
