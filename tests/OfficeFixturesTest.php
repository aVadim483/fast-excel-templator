<?php

declare(strict_types=1);

namespace avadim\FastExcelTemplator;

use PHPUnit\Framework\TestCase;

final class OfficeFixturesTest extends TestCase
{
    private const SHEET_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

    /** @dataProvider fixtures */
    public function testOfficeMetadataIsPreserved(string $name): void
    {
        $tpl = __DIR__ . '/fixtures/' . $name;
        $out = sys_get_temp_dir() . '/fxt-office-' . uniqid('', true) . '.xlsx';
        try {
            $excel = Excel::template($tpl, $out, ['strict_metadata' => true]);
            foreach ($excel->sheets() as $sheet) {
                $sheet->fill(['{{NAME}}' => 'Report'])->transferRows(null, static function ($source, $target, RowTemplate $row) {
                    if ($source === 2) {
                        $row->setValue('F', '=B2*2');
                    }
                    return $row;
                });
            }
            self::assertTrue($excel->save());
            self::assertSame([], $excel->metadataDiagnostics());
            $source = new \ZipArchive();
            $result = new \ZipArchive();
            self::assertTrue($source->open($tpl));
            self::assertTrue($result->open($out));
            try {
                $original = new \DOMDocument();
                $generated = new \DOMDocument();
                self::assertTrue($original->loadXML($source->getFromName('xl/worksheets/sheet1.xml')));
                self::assertTrue($generated->loadXML($result->getFromName('xl/worksheets/sheet1.xml')));
                $xpath = new \DOMXPath($generated);
                $xpath->registerNamespace('s', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
                self::assertSame('B2*2', $xpath->evaluate('string(//s:c[@r="F2"]/s:f)'));
                foreach (['sheetViews', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'tableParts', 'legacyDrawing'] as $tag) {
                    $sourceNodes = $original->getElementsByTagNameNS('http://schemas.openxmlformats.org/spreadsheetml/2006/main', $tag);
                    $targetNodes = $generated->getElementsByTagNameNS('http://schemas.openxmlformats.org/spreadsheetml/2006/main', $tag);
                    self::assertSame($sourceNodes->length, $targetNodes->length, $name . ': ' . $tag);
                    foreach ($sourceNodes as $index => $node) {
                        self::assertSame($node->C14N(true), $targetNodes->item($index)->C14N(true), $name . ': ' . $tag);
                    }
                }
                self::assertSame($source->getFromName('xl/workbook.xml'), $result->getFromName('xl/workbook.xml'));
                $original->loadXML($source->getFromName('xl/styles.xml'));
                $generated->loadXML($result->getFromName('xl/styles.xml'));
                $ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
                foreach (['dxf', 'tableStyles', 'colors'] as $tag) {
                    $sourceNodes = $original->getElementsByTagNameNS($ns, $tag);
                    $targetNodes = $generated->getElementsByTagNameNS($ns, $tag);
                    // Writer may add its default empty style sections or indexed palette.
                    if ($sourceNodes->length === 0) { continue; }
                    self::assertSame($sourceNodes->length, $targetNodes->length, $tag);
                    foreach ($sourceNodes as $index => $node) {
                        self::assertSame($node->C14N(true), $targetNodes->item($index)->C14N(true), $tag);
                    }
                }
            } finally { $source->close(); $result->close(); }
            self::assertSame('Report', \avadim\FastExcelReader\Excel::open($out)->readRows()[2]['A']);
        } finally { gc_collect_cycles(); @unlink($out); }
    }

    /** @dataProvider fixtures */
    public function testOfficeRulesFollowRowsAfterInsertsAndSkip(string $name): void
    {
        $tpl = __DIR__ . '/fixtures/' . $name;
        $out = sys_get_temp_dir() . '/fxt-office-rows-' . uniqid('', true) . '.xlsx';
        try {
            // The fixtures also contain defined names and other metadata outside the rule remapping scope.
            $excel = Excel::template($tpl, $out);
            $excel->sheet('Data')->transferRowsUntil(1)
                ->insertRow(['B' => 7])->skipRows(1)->insertRow(['B' => 8])->transferRows();
            self::assertTrue($excel->save());
            self::assertSame(
                ['unmapped-ranges', 'unmapped-defined-names'],
                array_column($excel->metadataDiagnostics(), 'code')
            );

            $original = $this->readWorksheet($tpl);
            $generated = $this->readWorksheet($out);
            foreach (['dataValidation', 'conditionalFormatting'] as $tag) {
                $sourceRules = $original->getElementsByTagNameNS(self::SHEET_NS, $tag);
                $outputRules = $generated->getElementsByTagNameNS(self::SHEET_NS, $tag);
                self::assertSame(1, $sourceRules->length, $tag);
                self::assertSame(1, $outputRules->length, $tag);
                $sourceRule = $sourceRules->item(0);
                $outputRule = $outputRules->item(0);
                self::assertSame('B2:B3', $sourceRule->getAttribute('sqref'), $tag);
                self::assertSame('B4', $outputRule->getAttribute('sqref'), $tag);
                $sourceRule->setAttribute('sqref', 'B4');
                self::assertSame($sourceRule->C14N(true), $outputRule->C14N(true), $tag);
            }

            $xpath = new \DOMXPath($generated);
            $xpath->registerNamespace('s', self::SHEET_NS);
            self::assertSame('1', $xpath->evaluate('string(//s:dataValidations/@count)'));
            self::assertSame('7', $xpath->evaluate('string(//s:c[@r="B2"]/s:v)'));
            self::assertSame('8', $xpath->evaluate('string(//s:c[@r="B3"]/s:v)'));
            self::assertSame('3', $xpath->evaluate('string(//s:c[@r="B4"]/s:v)'));
        } finally { gc_collect_cycles(); @unlink($out); }
    }

    private function readWorksheet(string $file): \DOMDocument
    {
        $zip = new \ZipArchive();
        self::assertTrue($zip->open($file));
        try {
            $dom = new \DOMDocument();
            self::assertTrue($dom->loadXML($zip->getFromName('xl/worksheets/sheet1.xml')));
            return $dom;
        } finally {
            $zip->close();
        }
    }

    public static function fixtures(): array
    {
        return [['metadata-excel.xlsx'], ['metadata-libreoffice.xlsx']];
    }
}
