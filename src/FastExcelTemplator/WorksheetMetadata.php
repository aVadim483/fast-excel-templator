<?php

declare(strict_types=1);

namespace avadim\FastExcelTemplator;

use avadim\FastExcelHelper\Helper;

/** @internal Stream regenerated row data while retaining template worksheet metadata. */
final class WorksheetMetadata
{
    private const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    private const BEFORE = ['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols'];
    private const AFTER = [
        'sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter', 'sortState',
        'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr', 'conditionalFormatting',
        'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter',
        'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags',
        'drawing', 'legacyDrawing', 'legacyDrawingHF', 'picture', 'oleObjects', 'controls',
        'webPublishItems', 'tableParts', 'extLst',
    ];
    private bool $strict;
    /** @var array<int, array{sheet: string, code: string, message: string}> */
    private array $diagnostics = [];

    public function __construct(bool $strict)
    {
        $this->strict = $strict;
    }

    public function diagnostics(): array
    {
        return $this->diagnostics;
    }

    /** Preserve auxiliary style sections whose indices are referenced by copied metadata. */
    public static function preserveStyleMetadata(string $source, string $generated): string
    {
        $original = new \DOMDocument();
        $result = new \DOMDocument();
        if (!$original->loadXML($source, LIBXML_NONET) || !$result->loadXML($generated, LIBXML_NONET)) {
            throw new Exception('Cannot preserve style metadata');
        }
        foreach (['tableStyles', 'colors', 'extLst'] as $name) {
            $node = self::rootChild($original, $name);
            if ($node !== null) {
                $existing = self::rootChild($result, $name);
                $copy = $result->importNode($node, true);
                if ($existing !== null) { $result->documentElement->replaceChild($copy, $existing); }
                else { $result->documentElement->appendChild($copy); }
            }
        }
        return $result->saveXML();
    }

    private static function rootChild(\DOMDocument $dom, string $name): ?\DOMElement
    {
        foreach ($dom->documentElement->childNodes as $child) {
            if ($child instanceof \DOMElement && $child->localName === $name && $child->namespaceURI === self::NS) {
                return $child;
            }
        }
        return null;
    }

    private function report(string $path, string $code, string $message): void
    {
        $this->diagnostics[] = ['sheet' => $path, 'code' => $code, 'message' => $message];
        if ($this->strict) {
            throw new Exception('Unsupported worksheet metadata in "' . $path . '": ' . $message);
        }
    }

    private function reader(string $file, string $path): \XMLReader
    {
        $reader = new \XMLReader();
        if (!$reader->open('zip://' . str_replace('\\', '/', $file) . '#' . $path, null, LIBXML_NONET)) {
            throw new Exception('Cannot read worksheet metadata "' . $path . '"');
        }
        return $reader;
    }

    private function attributes(\XMLReader $reader): array
    {
        $attributes = [];
        while ($reader->moveToNextAttribute()) {
            $attributes[$reader->name] = $reader->value;
        }
        $reader->moveToElement();
        return $attributes;
    }

    /** Only metadata is retained in memory; XMLReader::next() skips the source sheetData. */
    private function source(string $file, string $path): array
    {
        $reader = $this->reader($file, $path);
        $nodes = [];
        $attributes = [];
        try {
            $ready = $reader->read();
            while ($ready) {
                if ($reader->nodeType === \XMLReader::ELEMENT && $reader->depth === 0) {
                    if ($reader->localName !== 'worksheet' || $reader->namespaceURI !== self::NS) {
                        throw new Exception('Unsupported worksheet metadata namespace in "' . $path . '"');
                    }
                    $attributes = $this->attributes($reader);
                } elseif ($reader->nodeType === \XMLReader::ELEMENT && $reader->depth === 1) {
                    $name = $reader->localName;
                    if ($name !== 'sheetData') {
                        if ($reader->namespaceURI !== self::NS || !in_array($name, array_merge(self::BEFORE, self::AFTER), true)) {
                            $this->report($path, 'unsupported-node', 'Element ' . $reader->name . ' is not supported and will be omitted.');
                        } else {
                            $nodes[$name][] = $reader->readOuterXml();
                        }
                    }
                    $ready = $reader->next();
                    continue;
                }
                $ready = $reader->read();
            }
        } finally {
            $reader->close();
        }
        return [$attributes, $nodes];
    }

    /** @param resource $output */
    private function write($output, string $xml): void
    {
        $length = strlen($xml);
        $offset = 0;
        while ($offset < $length) {
            $written = fwrite($output, substr($xml, $offset));
            if ($written === false || $written === 0) {
                throw new Exception('Cannot write worksheet metadata');
            }
            $offset += $written;
        }
    }

    private function tag(string $name, array $attributes): string
    {
        $xml = '<' . $name;
        foreach ($attributes as $key => $value) {
            $xml .= ' ' . $key . '="' . htmlspecialchars($value, ENT_QUOTES | ENT_XML1, 'UTF-8') . '"';
        }
        return $xml . '>';
    }

    private function validations(string $source, string $generated): string
    {
        $dom = new \DOMDocument();
        $extra = new \DOMDocument();
        if (!$dom->loadXML($source, LIBXML_NONET) || !$extra->loadXML($generated, LIBXML_NONET)) {
            throw new Exception('Cannot merge data validation metadata');
        }
        foreach ($extra->documentElement->childNodes as $child) {
            if ($child instanceof \DOMElement) {
                $dom->documentElement->appendChild($dom->importNode($child, true));
            }
        }
        $count = 0;
        foreach ($dom->documentElement->childNodes as $child) {
            if ($child instanceof \DOMElement && $child->localName === 'dataValidation') { $count++; }
        }
        $dom->documentElement->setAttribute('count', (string)$count);
        return $dom->saveXML($dom->documentElement);
    }

    /** Map source rules before combining them with rules already in output coordinates. */
    private function remapRules(array $source, SheetTemplate $sheet): array
    {
        $mapping = $sheet->rowMapping();
        foreach (['dataValidations', 'conditionalFormatting'] as $kind) {
            foreach ($source[$kind] ?? [] as $index => $xml) {
                $dom = new \DOMDocument();
                if (!$dom->loadXML($xml, LIBXML_NONET)) {
                    throw new Exception('Cannot read rule metadata for "' . $sheet->name() . '"');
                }
                $root = $dom->documentElement;
                if ($kind === 'dataValidations') {
                    $count = 0;
                    foreach (iterator_to_array($root->childNodes) as $rule) {
                        if ($rule instanceof \DOMElement && $rule->namespaceURI === self::NS
                            && $rule->localName === 'dataValidation') {
                            if ($this->remapRule($rule, $mapping, $sheet)) {
                                $count++;
                            } else {
                                $root->removeChild($rule);
                            }
                        }
                    }
                    $root->setAttribute('count', (string)$count);
                    $keep = $count > 0;
                } else {
                    $keep = $this->remapRule($root, $mapping, $sheet);
                }
                if ($keep) {
                    $source[$kind][$index] = $dom->saveXML($root);
                } else {
                    unset($source[$kind][$index]);
                }
            }
            if (empty($source[$kind])) {
                unset($source[$kind]);
            }
        }
        return $source;
    }

    /** Return false when a rule no longer covers any source-derived row. */
    private function remapRule(\DOMElement $rule, RowMapping $mapping, SheetTemplate $sheet): bool
    {
        $sqref = $rule->getAttribute('sqref');
        $context = $rule->localName . ' on "' . $sheet->name() . '" (sqref="' . $sqref . '")';
        $areas = [];
        $anchor = null;
        $shift = null;
        $uniform = true;
        foreach (preg_split('/\s+/', trim($sqref)) as $area) {
            if (!preg_match('/^([A-Z]{1,3})([1-9][0-9]{0,6})(?::([A-Z]{1,3})([1-9][0-9]{0,6}))?$/iD', $area, $cell)) {
                $this->report($sheet->path(), 'unmapped-rule-range', $context . ' has an unsupported reference.');
                return true;
            }
            $first = (int)$cell[2];
            $last = (int)($cell[4] ?? $first);
            $left = strtoupper($cell[1]);
            $right = strtoupper($cell[3] ?? $left);
            if ($last < $first || $last > Helper::EXCEL_2007_MAX_ROW
                || Helper::colNumber($left) > Helper::colNumber($right)
                || Helper::colNumber($right) > Helper::EXCEL_2007_MAX_COL) {
                $this->report($sheet->path(), 'unmapped-rule-range', $context . ' has invalid bounds.');
                return true;
            }
            $anchor = $anchor ?? [$left, $first];
            foreach ($mapping->project($first, $last) as [$start, $end, $delta]) {
                $shift = $shift ?? ($start - $anchor[1]);
                $uniform = $uniform && $shift === $delta && ($areas || $anchor[0] === $left);
                $startCell = $left . $start;
                $endCell = $right . $end;
                $areas[] = $startCell === $endCell ? $startCell : $startCell . ':' . $endCell;
            }
        }
        if (!$areas) {
            return false;
        }
        $mapped = implode(' ', $areas);
        if ($mapped !== $sqref || !$uniform) {
            $this->remapRuleFormulas($rule, $uniform ? $shift : null, $sheet, $context);
            $rule->setAttribute('sqref', $mapped);
        }
        return true;
    }

    private function remapRuleFormulas(\DOMElement $rule, ?int $shift, SheetTemplate $sheet, string $context): void
    {
        $formulas = [];
        foreach (['formula1', 'formula2', 'formula', 'cfvo'] as $tag) {
            foreach ($rule->getElementsByTagNameNS(self::NS, $tag) as $node) {
                if ($tag === 'cfvo' && $node->getAttribute('type') !== 'formula') {
                    continue;
                }
                $attribute = $tag === 'cfvo' ? $node->getAttributeNode('val') : null;
                $formulas[] = $attribute ?: $node;
            }
        }
        $replacements = [];
        foreach ($formulas as $node) {
            $formula = $this->shiftRuleFormula($node->nodeValue, $shift);
            if ($formula === null) {
                $this->report($sheet->path(), 'unmapped-rule-formula', $context . ' has a formula whose meaning cannot be preserved after remapping.');
                return;
            }
            $replacements[] = $formula;
        }
        foreach ($formulas as $index => $node) {
            if ($node->nodeValue !== $replacements[$index]) {
                if ($node instanceof \DOMAttr) {
                    $node->ownerElement->setAttribute($node->name, $replacements[$index]);
                } else {
                    $node->textContent = $replacements[$index];
                }
            }
        }
    }

    /**
     * Translate ordinary relative row references only when the entire rule moves uniformly.
     * Context-dependent functions, named ranges, sheet prefixes and other unverified syntax
     * must be diagnosed. Scalar values and formulae containing only fixed rows need no shift.
     */
    private function shiftRuleFormula(string $formula, ?int $shift): ?string
    {
        if ($shift === 0) {
            return $formula;
        }
        $syntax = FormulaLexer::replaceReferences($formula, static function (string $ref): string {
            return '0';
        });
        $syntax = preg_replace('/"(?:[^"]|"")*"/', '0', $syntax);
        $allowedSyntax = '~^(?:[\s0-9.+*/^%&<>=(),-]|\b(?:TRUE|FALSE)\b'
            . '|\b(?:AND|OR|NOT)(?=\s*\()|(?<=\d)[Ee](?=[+-]?\d))*$~iD';
        if (!preg_match($allowedSyntax, $syntax)) {
            return null;
        }
        $valid = true;
        $result = FormulaLexer::replaceReferences($formula, static function (string $ref) use ($shift, &$valid): string {
            return preg_replace_callback('/(\$?[A-Z]{1,3})(\$?)([0-9]+)/i', static function (array $cell) use ($shift, &$valid): string {
                if ($cell[2] === '$') {
                    return $cell[0];
                }
                $row = (int)$cell[3] + $shift;
                if ($shift === null || $row < 1 || $row > Helper::EXCEL_2007_MAX_ROW) {
                    $valid = false;
                    return $cell[0];
                }
                return $cell[1] . $row;
            }, $ref);
        });
        return $valid ? $result : null;
    }

    /** @param resource $output */
    private function metadata($output, array $order, array $source, array $generated, string $path): void
    {
        foreach ($order as $name) {
            $nodes = $source[$name] ?? $generated[$name] ?? [];
            // Coordinates, widths, row format and merges come from the generated rows.
            if (in_array($name, ['dimension', 'cols', 'mergeCells'], true)) {
                $nodes = $generated[$name] ?? [];
            } elseif ($name === 'dataValidations' && isset($source[$name][0], $generated[$name][0])) {
                $nodes = [$this->validations($source[$name][0], $generated[$name][0])];
            } elseif ($name === 'conditionalFormatting' && isset($source[$name], $generated[$name])) {
                $this->report($path, 'rule-conflict', 'Template and generated conditional rules require explicit priority reconciliation.');
                $nodes = array_merge($source[$name], $generated[$name]);
            } elseif ($name === 'extLst' && isset($source[$name], $generated[$name])) {
                $this->report($path, 'extension-conflict', 'Template extensions take precedence over newly generated extensions; automatic reconciliation is not supported.');
            } elseif (in_array($name, ['autoFilter', 'pageMargins', 'pageSetup', 'headerFooter', 'sheetFormatPr'], true)
                && isset($source[$name][0], $generated[$name][0])) {
                // Preserve original child nodes and attributes; explicitly changed attributes win.
                $original = new \DOMDocument();
                $new = new \DOMDocument();
                $original->loadXML($source[$name][0], LIBXML_NONET);
                $new->loadXML($generated[$name][0], LIBXML_NONET);
                foreach ($new->documentElement->attributes as $attribute) {
                    if ($attribute->namespaceURI === null) {
                        $original->documentElement->setAttribute($attribute->name, $attribute->value);
                    }
                }
                if ($name === 'headerFooter') {
                    foreach ($new->documentElement->childNodes as $child) {
                        if (!$child instanceof \DOMElement) { continue; }
                        $old = $original->getElementsByTagNameNS(self::NS, $child->localName)->item(0);
                        $copy = $original->importNode($child, true);
                        if ($old !== null) { $original->documentElement->replaceChild($copy, $old); }
                        else { $original->documentElement->appendChild($copy); }
                    }
                }
                $nodes = [$original->saveXML($original->documentElement)];
            }
            foreach ($nodes as $xml) { $this->write($output, $xml); }
        }
    }

    /** Return a temporary sheet XML, to be retained until ZipArchive::close(). */
    public function rewrite(string $input, string $generated, SheetTemplate $sheet, string $directory): string
    {
        $path = $sheet->path();
        [$rootAttributes, $source] = $this->source($input, $path);
        if ($sheet->hasStructuralChanges()) {
            if (isset($source['dataValidations']) || isset($source['conditionalFormatting'])) {
                $unmapped = $sheet->unmappedRowReason();
                if ($unmapped === null) {
                    $source = $this->remapRules($source, $sheet);
                } else {
                    $this->report($path, 'unmapped-rule-ranges', 'Rules on "' . $sheet->name() . '": ' . $unmapped);
                }
            }
            $sensitive = ['autoFilter', 'hyperlinks', 'tableParts',
                'drawing', 'legacyDrawing', 'extLst', 'sortState', 'rowBreaks', 'colBreaks'];
            $unmapped = array_intersect($sensitive, array_keys($source));
            if ($unmapped) {
                $this->report($path, 'unmapped-ranges', 'Template coordinates are retained after structural changes for: ' . implode(', ', $unmapped) . '.');
            }
            // Defined names are kept in workbook.xml, including print areas and named ranges.
            $workbook = $this->reader($input, 'xl/workbook.xml');
            try {
                while ($workbook->read()) {
                    if ($workbook->nodeType === \XMLReader::ELEMENT && $workbook->localName === 'definedName') {
                        $this->report($path, 'unmapped-defined-names', 'Workbook defined names retain template coordinates after structural changes.');
                        break;
                    }
                }
            } finally { $workbook->close(); }
        }
        $temporary = tempnam($directory, 'fxt-meta-');
        if ($temporary === false) { throw new Exception('Cannot create metadata output'); }
        $reader = null;
        $output = null;
        $success = false;
        try {
            $reader = $this->reader($generated, $path);
            $output = fopen($temporary, 'wb');
            if ($output === false) { throw new Exception('Cannot open metadata output'); }
            $nodes = [];
            $ready = $reader->read();
            while ($ready) {
                if ($reader->nodeType === \XMLReader::ELEMENT && $reader->depth === 0) {
                    $attributes = array_merge($this->attributes($reader), $rootAttributes);
                    $this->write($output, '<?xml version="1.0" encoding="UTF-8"?>' . $this->tag('worksheet', $attributes));
                } elseif ($reader->nodeType === \XMLReader::ELEMENT && $reader->depth === 1) {
                    if ($reader->localName === 'sheetData') {
                        $this->metadata($output, self::BEFORE, $source, $nodes, $path);
                        $this->write($output, '<sheetData>');
                        if (!$reader->isEmptyElement) {
                            $ready = $reader->read();
                            while ($ready && !($reader->depth === 1 && $reader->nodeType === \XMLReader::END_ELEMENT)) {
                                if ($reader->depth === 2 && $reader->nodeType === \XMLReader::ELEMENT) {
                                    $this->write($output, $reader->readOuterXml());
                                    $ready = $reader->next();
                                } else { $ready = $reader->read(); }
                            }
                        }
                        $this->write($output, '</sheetData>');
                    } else {
                        $nodes[$reader->localName][] = $reader->readOuterXml();
                        $ready = $reader->next();
                        continue;
                    }
                }
                $ready = $reader->read();
            }
            $this->metadata($output, self::AFTER, $source, $nodes, $path);
            $this->write($output, '</worksheet>');
            if (!fflush($output)) { throw new Exception('Cannot flush metadata output'); }
            $closed = fclose($output);
            $output = null;
            if (!$closed) { throw new Exception('Cannot close metadata output'); }
            $success = true;
            return $temporary;
        } finally {
            if ($reader !== null) { $reader->close(); }
            if (is_resource($output)) { fclose($output); }
            if (!$success) { @unlink($temporary); }
        }
    }
}
