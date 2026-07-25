# Changelog

🇬🇧 English · [🇷🇺 Русский](CHANGELOG.ru.md)

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
