<?php

declare(strict_types=1);

namespace avadim\FastExcelTemplator;

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
    public function rewrite(string $input, string $generated, string $path, bool $changed, string $directory): string
    {
        [$rootAttributes, $source] = $this->source($input, $path);
        if ($changed) {
            $sensitive = ['autoFilter', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'tableParts',
                'drawing', 'legacyDrawing', 'extLst', 'sortState', 'rowBreaks', 'colBreaks'];
            if (array_intersect($sensitive, array_keys($source))) {
                $this->report($path, 'unmapped-ranges', 'Metadata ranges and drawing anchors retain template coordinates after structural changes.');
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
