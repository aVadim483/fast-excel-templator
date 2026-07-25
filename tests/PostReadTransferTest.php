<?php

declare(strict_types=1);

namespace avadim\FastExcelTemplator;

use avadim\FastExcelWriter\Excel as Writer;
use avadim\FastExcelWriter\DataValidation\DataValidation;
use PHPUnit\Framework\TestCase;

/**
 * Regression tests for sheet-level nodes transferred after </sheetData> (SheetTemplate::postRead()):
 *  - issue #21: the autofilter range from the template must be preserved
 *  - issue #2:  data validations (dropdown lists) must be transferred
 */
final class PostReadTransferTest extends TestCase
{
    private function sheetXml(string $file): string
    {
        $zip = new \ZipArchive();
        self::assertTrue($zip->open($file) === true);
        $xml = $zip->getFromName('xl/worksheets/sheet1.xml');
        $zip->close();

        return (string)$xml;
    }

    public function testAutoFilterAndDataValidationAreTransferred()
    {
        $tpl = __DIR__ . '/tpl-postread.xlsx';
        $out = __DIR__ . '/out-postread.xlsx';
        @unlink($tpl);
        @unlink($out);

        // Build a template that contains an autofilter and a dropdown (data validation)
        $writer = Writer::create(['Sheet1']);
        $ws = $writer->sheet();
        $ws->writeRow(['Name', 'Qty', 'Choice']);
        $ws->writeRow(['{{NAME}}', 0, '']);
        $ws->writeRow(['b', 1, '']);
        $ws->writeRow(['c', 2, '']);
        $ws->setAutoFilter('A1:C4');
        $ws->addDataValidation('C2:C100', DataValidation::dropDown(['Yes', 'No']));
        $writer->save($tpl);

        // Run the template through the templator
        $excel = Excel::template($tpl, $out);
        $sheet = $excel->sheet();
        $sheet->fill(['{{NAME}}' => 'ACME']);
        $sheet->transferRows();
        $excel->save();

        $xml = $this->sheetXml($out);

        // issue #21: the autofilter range from the template must survive (not a hardcoded A1, not dropped)
        self::assertStringContainsString('<autoFilter ref="A1:C4"/>', $xml);

        // issue #2: the dropdown (data validation) must be transferred with its range and list values
        self::assertStringContainsString('<dataValidation', $xml);
        self::assertStringContainsString('type="list"', $xml);
        self::assertStringContainsString('sqref="C2:C100"', $xml);
        self::assertStringContainsString('Yes,No', $xml);

        @unlink($tpl);
        @unlink($out);
    }
}
