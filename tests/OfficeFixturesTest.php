<?php

declare(strict_types=1);

namespace avadim\FastExcelTemplator;

use PHPUnit\Framework\TestCase;

final class OfficeFixturesTest extends TestCase
{
    /** @dataProvider fixtures */
    public function testOfficeMetadataIsPreserved(string $name): void
    {
        $tpl = __DIR__ . '/fixtures/' . $name;
        $out = sys_get_temp_dir() . '/fxt-office-' . uniqid('', true) . '.xlsx';
        try {
            $excel = Excel::template($tpl, $out, ['strict_metadata' => true]);
            foreach ($excel->sheets() as $sheet) {
                $sheet->fill(['{{NAME}}' => 'Report'])->transferRows();
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

    public static function fixtures(): array
    {
        return [['metadata-excel.xlsx'], ['metadata-libreoffice.xlsx']];
    }
}
