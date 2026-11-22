# XLSX feature preservation

English · [Русский](XLSX_SUPPORT.ru.md)

XLSX using the SpreadsheetML Transitional namespace is supported. The template ZIP is copied, sheet data is regenerated, and worksheet metadata is retained in a separate streaming pass. The pass holds metadata, compact row mappings and one XML row in memory. Referenced template archive parts remain in place.

| Feature | Replacing values without structural changes | Inserting/skipping rows or shifting cells |
|---|---|---|
| Conditional formatting | Original rules, ranges, priorities and `dxf` indices are retained | Standard `sqref` ranges follow tracked row changes. Formula changes are subject to the rules below |
| Data validation | Attributes, formulas, child nodes, container options and original extensions are retained | Standard `sqref` ranges are remapped before Writer validations are appended. Empty rules are removed and the count is updated |
| Autofilter | Original range and filter children survive; generated Writer attributes take precedence | No automatic expansion; reported, rejected in strict mode |
| Excel tables | Original `tableParts`, table XML, styles and relationships survive | Table ranges, filters and calculated columns are not rebuilt; reported, rejected in strict mode |
| Hyperlinks | Original nodes and relationships survive, including external targets | Cell addresses are not remapped; reported, rejected in strict mode |
| Defined names and print areas | Original workbook XML survives | Defined-name formulas remain unchanged; reported, rejected in strict mode |
| Images and notes | Original media, drawings, VML, comments and relationships survive | Anchors and note addresses remain unchanged; reported, rejected in strict mode |
| Print settings | Original nodes survive; generated Writer attributes take precedence. Header/footer text and XML false flags are retained | Print areas and breaks are not remapped |
| Sheet views | Original selections and frozen/split panes survive | Selection and pane coordinates are not remapped |
| Merged cells | Recreated from source rows | Horizontal merges shift/shrink with removed cells; removing their anchor removes the merge. Cell removals affecting a multi-row merge throw an exception |
| Palettes and table styles | Original `colors`, `tableStyles` and style extensions survive | No structural remapping |
| `extLst` extensions | Original XML and namespaces survive | Extension semantics are not interpreted; structural changes are reported |
| Unknown top-level elements | Omitted with a diagnostic | Strict mode rejects the save |

Combining original and newly generated conditional rules requires explicit priority reconciliation. When both template and generated `extLst` exist, the template extension takes precedence. Both conflicts produce diagnostics and are rejected in strict mode.

## How rule ranges follow rows

Single cells, rectangles and lists of areas are handled per sheet. For example, when source row 2 is skipped, `B3:B5` becomes `B2:B4`. Each copy of a row template from that sheet inherits the rules covering the source row. A blank row inserted without a source template has no inherited coverage and can split an area. Sparse positions keep their spacing. When a callback stops a transfer, the stopped row is consumed. Unread rows enter the mapping if a later transfer reaches them.

A rule formula based on ordinary A1 references and arithmetic or comparison operators follows a uniform row shift. References with absolute row numbers and scalar formulas such as validation bounds or a literal list retain their text. A rule with a relative formula and mixed row displacements is diagnosed, as are formula syntaxes whose anchoring cannot be verified. Default mode applies the new `sqref` and retains the original formula in those cases. Strict mode rejects the save.

## Diagnostics and strict mode

```php
$excel = Excel::template($template, $output, ['strict_metadata' => true]);
$excel->sheet()->fill(['{{NAME}}' => 'Report'])->transferRows();
$excel->save();
$diagnostics = $excel->metadataDiagnostics();
```

Diagnostics contain `sheet`, `code` and `message`, populated during the latest archive-generation attempt. `sheet` is an XML part path such as `xl/worksheets/sheet1.xml`; messages are in English. Preliminary path-validation failures do not refresh metadata diagnostics.

`strict_metadata` defaults to `false`: saving proceeds with the documented limitations. Strict mode throws on a detected unsupported transformation before publishing the result, preserving an existing destination. It stops at the first problem, so its report may be incomplete.

| Code | Meaning | Default-mode behavior |
|---|---|---|
| `unsupported-node` | Unknown top-level worksheet element | Omit the element |
| `unmapped-ranges` | Structural changes without metadata range/anchor remapping | Retain template coordinates |
| `unmapped-rule-ranges` | Untracked Writer writes, direct source reads, or cells shifted horizontally | Retain template rule coordinates |
| `unmapped-rule-range` | Invalid or unsupported `sqref` | Retain that rule's coordinates |
| `unmapped-rule-formula` | Formula semantics cannot be verified after an area changes | Keep the formula text and apply the remapped area |
| `unmapped-defined-names` | Structural changes in a workbook containing defined names | Retain defined-name formulas |
| `rule-conflict` | Both template and generated conditional rules exist | Keep both without reconciling priorities |
| `extension-conflict` | Both template and generated `extLst` exist | Template `extLst` takes precedence |

Handling rejection:

```php
use avadim\FastExcelTemplator\Exception;

try {
    $excel->save();
} catch (Exception $exception) {
    foreach ($excel->metadataDiagnostics() as $item) {
        error_log($item['sheet'] . ': ' . $item['code'] . ' — ' . $item['message']);
    }
    throw $exception;
}
```

`SheetTemplate::hasStructuralChanges()` reports whether source coordinates may have changed. Metadata outside the supported rule ranges is still checked conservatively, and defined names are checked workbook-wide. Detected low-level Writer row or cell writes are treated as untracked changes.

## Values, links and development

`RowTemplate::setValue('A', null)` clears content and the old formula while retaining formatting. Zero, `false` and the empty string remain distinct valid values.

`removeCell()`/`removeCells()` shift cells in the current row; they do not structurally delete a workbook column or rewrite references throughout the workbook. A moved formula rebases relative references at its new address and retains absolute references. References in other cells are not rewritten.

Whole-row, whole-column and spilled-range references in cell formulas retain their original text.

Relative paths, Windows case aliases and symbolic links resolving to the template are rejected as output paths. Replacing a destination hard link replaces only that directory entry, leaving the template intact. Exclusive saves reject existing links, including dangling symlinks.

`tests/fixtures/` contains Excel and LibreOffice workbooks. Tests inspect output values and formulas, rule XML, node order, styles, ZIP relationships, images, notes and reverse sheet-processing order. Office applications are not required by PHPUnit.

Run `composer test` and `composer analyse`. CI covers Windows/Ubuntu and PHP 7.4, 8.0–8.6. Lowest dependencies are checked separately on PHP 7.4; other jobs install current compatible versions. PHPStan level 1 targets PHP 7.4 without baseline entries or suppressed errors. The new internal metadata handler uses strict types; existing API coercion rules are retained.

The local `TemplateWriter` adapter checks archive open, entry replacement and close results for the template save path.
