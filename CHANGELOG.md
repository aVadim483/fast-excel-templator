# Changelog

🇬🇧 English · [🇷🇺 Русский](CHANGELOG.ru.md)

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
