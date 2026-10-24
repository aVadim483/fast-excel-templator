# XLSX feature preservation

English · [Русский](XLSX_SUPPORT.ru.md)

XLSX using the SpreadsheetML Transitional namespace is supported. The template ZIP is copied, sheet data is regenerated, and worksheet metadata is retained in a separate streaming pass. Only metadata and one XML row are held in memory. Referenced template archive parts remain in place.

| Feature | Replacing values without structural changes | Inserting/skipping rows or shifting cells |
|---|---|---|
| Conditional formatting | Original rules, ranges, priorities and `dxf` indices are retained; differential styles keep their indices | Rule ranges and formulas retain template coordinates; reported, rejected in strict mode |
| Data validation | All attributes, formulas, child nodes, container options and original extensions survive. New ordinary Writer validations are appended | Ranges and formulas are not remapped; reported, rejected in strict mode |
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

`SheetTemplate::hasStructuralChanges()` exposes the conservative structural-operation flag. Insertions, skips and removals can trigger rejection even when a particular range is unaffected. Defined names are checked workbook-wide. Direct low-level Writer changes are not fully tracked. Strict mode does not validate the semantics of arbitrary third-party extensions.

## Values, links and development

`RowTemplate::setValue('A', null)` clears content and the old formula while retaining formatting. Zero, `false` and the empty string remain distinct valid values.

`removeCell()`/`removeCells()` shift cells in the current row; they do not structurally delete a workbook column or rewrite references throughout the workbook. A moved formula rebases relative references at its new address and retains absolute references. References in other cells are not rewritten.

Relative paths, Windows case aliases and symbolic links resolving to the template are rejected as output paths. Replacing a destination hard link replaces only that directory entry, leaving the template intact. Exclusive saves reject existing links, including dangling symlinks.

`tests/fixtures/` contains Excel and LibreOffice workbooks. Tests inspect XML, node order, styles, ZIP relationships, images, notes and reverse sheet-processing order. Office applications are not required by PHPUnit.

Run `composer test` and `composer analyse`. CI covers Windows/Ubuntu and PHP 7.4, 8.0–8.5. Lowest dependencies are checked separately on PHP 7.4; other jobs install current compatible versions. PHPStan level 1 targets PHP 7.4 without baseline entries or suppressed errors. The new internal metadata handler uses strict types; existing API coercion rules are retained.

ZIP checks remain in the local `TemplateWriter` adapter because the minimum Writer 6.16.1 does not provide them. Moving them upstream requires checking archive open, entry replacement and close results, closing on exceptions, and propagating failure. This implementation requires no dependency bump and does not patch `vendor/`.
