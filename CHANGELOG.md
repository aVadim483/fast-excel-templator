# Changelog

🇬🇧 English · [🇷🇺 Русский](CHANGELOG.ru.md)

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Fixed
- The template overwrite guard explicitly resolves file symlinks with `readlink()`, including link chains: Windows `realpath()` alone is insufficient.
- `insertRow()` uses a template copy: partial data no longer inherits values from a previous insertion. Collection selection for insertion is independent of `foreach`.
- `transferRows(0)` and `skipRows(0)` leave cursors unchanged; negative counts and limits throw an exception. Processing up to a sparse-sheet boundary no longer consumes the next existing row.
- Missing rows in a template range become empty templates; a self-closing XML row no longer takes cells from the following row. Invalid ranges and empty collections produce explicit exceptions.
- A separate streaming pass preserves worksheet metadata: conditional rules, complete validation attributes/children, hyperlinks, tables, drawing/comment references, views and extensions. Differential styles are loaded correctly; palettes, table styles and style extensions survive.
- The streaming path in newer Readers now rebases moved formulas. Horizontal merges shrink when interior cells are removed; removals affecting multi-row merges are rejected.
- `setValue(..., null)` clears cell content and the old formula while retaining formatting. Header/footer XML false flags and escaped text are handled correctly.
- Generated sheets are mapped to their original XML parts even when workbook order differs from worksheet part numbering.

### Added
- `strict_metadata` and `metadataDiagnostics()`: unsupported range transformations and unknown elements produce diagnostics; strict mode rejects saving before publication.
- Bilingual XLSX support matrices and fixtures created by Excel and LibreOffice.
- Windows/Ubuntu CI for PHP 7.4, 8.0–8.5, separate lowest-dependency jobs on PHP 7.4, and PHPStan level 1 without baseline entries or suppressed errors (`composer analyse`).

### Tests
- Added metadata, ZIP relationships, images, notes, filesystem links, null clearing, sparse rows, merge and Office fixture checks.
- Symlink tests use native Windows separators and verify the target before saving. The dangling-link test creates an existing target and removes it after creating the link, allowing it to run on Windows too.

### Documentation
- README and API reference now describe independent insertions, cursor rules and cell clearing. The support matrix documents diagnostic codes, strict rejection handling and structural limitations.

### Compatibility
- Callers relying on insertion mutating a template, or sharing an insertion cursor with `next()`/`foreach`, must supply values and select templates explicitly. Public `next()` behavior is preserved; independent `nextTemplate()` selection was added.
- Metadata ranges and drawing anchors retain template coordinates after structural changes. Strict mode conservatively rejects these operations; see XLSX_SUPPORT.md.

## [3.3.1] - 2026-10-08

### Fixed
- Saving generates a temporary file in the destination directory before publishing the completed workbook. Generation and publication failures preserve existing output; saving over the source template is rejected.
- `save(..., false)` refuses to overwrite, including a destination created during generation.
- Archive replacement checks ZIP open, entry write and close results and releases handles on exceptions. A local Writer adapter provides these checks for supported dependency versions.
- The templator writer no longer auto-saves on destruction. Failed downloads clean up their temporary files.
- `RowTemplate::setValue()` removes the previous formula and resets source type metadata while preserving styles.
- Removing cells moves styles and merged ranges to the same target address as the value.

### Tests
- Regression tests cover safe saving, ZIP failures, download cleanup, formula replacement and styles/merges after cell removal.

## [3.3.0] - 2026-08-29

### Fixed
- Two defects of `avadim/fast-excel-writer` that reached this package are gone with the writer 6.16.1:
  - a reference shifted off the sheet kept its internal RC text (`R[-4]C[-1]`) instead of becoming `#REF!`, and a range could come out half in RC and half in A1 (`SUM(R[-4]C[-1]:A2)`). Row templates keep their formulas in RC to re-base them on insert, so this reached every template inserted above its source row;
  - a workbook borrowed the locale of another workbook alive at the same time (a `ru` workbook wrote the euro sign), because the style manager kept a static instance overwritten by every constructor. The templator reached it through `RowTemplate::setCellStyle()` / `setRowStyle()`.
- `Excel::download()` announced the file as the old `.xls` (`application/vnd.ms-excel`) instead of the XLSX MIME type, and sent no `Content-Length`. It also wrote its headers into a response that had already started, appending a workbook to whatever was printed before; it now reports that instead.
- `SheetTemplate::readRow()` ended its loop on any falsy generator key rather than on the end of the generator — row numbers start at 1, so it worked by luck.
- Loading a `styles.xml` that holds an `<xf>` outside of any known section raised a `TypeError` instead of ignoring the element.

### Added
- `Excel::MIME_TYPE`, and the `downloadHeaders()` / `sendHeaders()` / `headersSent()` hooks behind `download()`, so a subclass can change what is sent (and a test can observe it).

### Changed
- Raised requirement: `avadim/fast-excel-writer` to `^6.16.1`.

### Tests
- The test held back by the writer (a reference shifted off the sheet must become `#REF!`) is no longer skipped, and a range shifted the same way is covered next to it.
- Two workbooks of different locales, alive at the same time, are asserted to keep their own currency.
- New coverage for `download()` / `output()`, `SheetTemplate::replaceRow()`, `cloneRow()`, `writeRow()`, `writeCell()`, `RowTemplate::removeCell()` / `removeCells()`, and workbooks of several sheets.
- The files the test suite writes are ignored by git instead of being committed.

## [3.2.0] - 2026-08-16

### Fixed
- Formulas are parsed by a dedicated lexer instead of the PHP tokenizer. The tokenizer read Excel formulas by PHP rules, so `#` started a comment and every reference after an error constant (`=SUM(#REF!)+A3`) was left un-rebased; a range fell apart into its two ends, which is how mixed `R[-4]C[-1]:A2` notation could reach a saved file; and `$A$1` was read as a variable. Ranges are now converted as a whole, and sheet prefixes, string literals, table references and numbers such as `1E3` are told apart from real references.
- A row was truncated when one of its cells held a falsy value: `RowTemplate` checked the current value instead of the key, so iteration stopped mid-row and every following cell was silently dropped.
- The relationship to the removed `calcChain.xml` was left dangling in the saved workbook — the part name was misspelled (`workbook.xml.res` instead of `.rels`).
- `Excel::outputFile()` threw an `Error` when `template()` had been called without an output file name.
- A date cell was written twice, the second time without its number format; a cell built by hand without the `t` key raised a warning.
- `SheetTemplate::mergedRange()` no longer overrides the inherited method with something else: it used to answer only for the top-left cell of a merge and return relative offsets instead of an address range. The offsets are internal and now live in their own method.
- `RowTemplateCollection` implements `Iterator` but could not be iterated — `next()` wraps around by design, so a `foreach` never ended. The collection also required a separate `setSheet()` call before `cloneCell()`, otherwise it threw on an uninitialized property; the sheet is now an optional constructor argument.
- `Excel::template()` returned `new self`, handing an `Excel` back to subclasses; it now returns `new static`.
- `Reader::validate()` echoed the file list with `<br>` tags instead of returning it, and threw when called before any document was open.
- `SheetWriter::getSheetViews()` reshaped the object's state while answering; the reshaping happens on input now.
- Errors are no longer swallowed: `<dataValidations>` transfer catches the writer's own validation exception rather than every `\Throwable`, and `save()` reports a file it could not overwrite instead of silencing the failure.

### Changed
- Raised requirements: `avadim/fast-excel-reader` to `^4.4`, `avadim/fast-excel-writer` to `^6.16`.
- Dropped the explicit `avadim/fast-excel-helper` requirement — it comes in through the reader and the writer, both of which require `^1.4`.
- Row range errors are reported with the package `Exception` instead of a bare `\RuntimeException` (it extends `RuntimeException`, so existing `catch` blocks keep working).
- `SheetTemplate::replaceRow()` and `cloneRow()` return `$this`, like the rest of the fluent API.

### Deprecated
- `Excel::$instance` — assigned but never read, and overwritten by every new workbook. Keep the object returned by `template()` instead. To be removed in 4.0.

### Performance
- Capturing a row template no longer scans every merge of the sheet for every cell of that row.

### Docs
- Added `BACKLOG.md` with the known issues left undone, including two bugs of `fast-excel-writer` that affect this package.
- Regenerated the API reference (it had been missing the reader 4.4 methods).

## [3.1.0] - 2026-07-25

### Changed
- Raised the `avadim/fast-excel-reader` requirement to `^4.1`, which renders built-in date formats deterministically regardless of the runtime locale and of whether `ext-intl` is loaded (ref reader #53).

## [3.0.1] - 2026-07-25

### Fixed
- Template `autoFilter` was lost and its range hardcoded to `A1`: the first sheet node after `</sheetData>` was skipped, so `<autoFilter>` (which comes first) was never transferred. The original filter range from the template is now preserved (#21).
- Data validations (e.g. dropdown lists) were not transferred to the output. `<dataValidations>` are now carried over from the template (#2).

## [3.0.0] - 2026-07-25

### Breaking
- Requires the new major of the base reader: `avadim/fast-excel-reader` `^4.0` (previously `^3`). Update your dependencies when upgrading.

### Changed
- Bumped `avadim/fast-excel-writer` to `^6.15`.
- `SheetTemplate::getRowTemplates()` now validates the requested row range against the sheet `<dimension>` and throws `RuntimeException` for out-of-range rows.

### Performance
- Merged cells are read lazily, avoiding a full scan at open time for sheets that are never transferred.
- The source DOM node is retained only for error cells instead of for every cell.

### Docs
- README overhaul (EN + RU): added "Which library do I need?", "How it works", "Limitations" and FAQ;
  documented multiple sheets, browser output (`download()`/`output()`), row-template ranges
  and the `fill()` vs `replace()` distinction; fixed the `replace()` example and the `rows()` method name.
- Enriched PHPDoc and regenerated the API reference; added the `RowTemplateCollection` reference page.
- Added this CHANGELOG (EN + RU) and a GitHub Actions CI workflow (PHP 7.4–8.4).

---

For releases prior to this changelog, see the
[GitHub Releases page](https://github.com/aVadim483/fast-excel-templator/releases).
