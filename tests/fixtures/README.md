# Office metadata fixtures

`metadata-excel.xlsx` was created with Microsoft Excel on 2026-10-24. It contains a conditional rule with a differential style, whole-number validation with input/error messages, a hyperlink, a note, a table, a named range, a print area and a frozen header row.

`metadata-libreoffice.xlsx` is the same workbook opened and saved as XLSX by LibreOffice on that date. It exercises the XML representation produced by Calc, including its extensions.

The fixtures contain synthetic data only. `OfficeFixturesTest` checks cell replacements, semantic XML equality of preserved metadata, and the rule ranges after rows are inserted and skipped. External Office applications are not required to run PHPUnit.
