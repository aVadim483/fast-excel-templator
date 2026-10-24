<?php

declare(strict_types=1);

namespace avadim\FastExcelTemplator;

use avadim\FastExcelWriter\Excel as Writer;
use PHPUnit\Framework\TestCase;

final class MetadataPreservationTest extends TestCase
{
    private string $tpl;
    private string $out;

    protected function setUp(): void
    {
        $this->tpl = sys_get_temp_dir() . '/fxt-metadata-' . uniqid('', true) . '.xlsx';
        $this->out = $this->tpl . '-out.xlsx';
        $writer = Writer::create(['Alpha', 'Beta']);
        foreach (['Alpha', 'Beta'] as $name) {
            $writer->sheet($name)->writeRow(['{{NAME}}', 2]);
            $writer->sheet($name)->writeRow(['value', 3]);
        }
        $writer->save($this->tpl);
        $zip = new \ZipArchive();
        self::assertTrue($zip->open($this->tpl));
        $styles = $zip->getFromName('xl/styles.xml');
        $styles = preg_replace('/<dxfs\b[^>]*(?:\/>|>.*?<\/dxfs>)/s',
            '<dxfs count="1"><dxf><font><b/><color rgb="FFFF0000"/></font><numFmt numFmtId="165" formatCode="0.000"/></dxf></dxfs>', $styles);
        self::assertTrue($zip->addFromString('xl/styles.xml', $styles));
        for ($i = 1; $i <= 2; $i++) {
            $entry = 'xl/worksheets/sheet' . $i . '.xml';
            $xml = $zip->getFromName($entry);
            $metadata = '<conditionalFormatting sqref="B1:B2"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>1</formula></cfRule></conditionalFormatting>'
                . '<dataValidations count="1" disablePrompts="1" xWindow="42" yWindow="17"><dataValidation type="whole" operator="between" sqref="B1 B2" allowBlank="true" showErrorMessage="1" errorTitle="Invalid &amp; value" error="Choose 1–9" showInputMessage="1" promptTitle="Hint" prompt="Enter a number" errorStyle="stop"><formula1>1</formula1><formula2>9</formula2></dataValidation></dataValidations>'
                . '<hyperlinks><hyperlink ref="A1" location="Beta!A1" display="Go" tooltip="Internal link"/></hyperlinks>';
            $xml = str_replace('</sheetData>', '</sheetData>' . $metadata, $xml);
            self::assertTrue($zip->addFromString($entry, $xml));
        }
        self::assertTrue($zip->close());
    }

    protected function tearDown(): void
    {
        gc_collect_cycles();
        @unlink($this->tpl);
        @unlink($this->out);
    }

    private function part(string $file, string $entry): string
    {
        $zip = new \ZipArchive();
        self::assertTrue($zip->open($file));
        $xml = $zip->getFromName($entry);
        $zip->close();
        self::assertIsString($xml);
        return $xml;
    }

    private function xpath(string $xml): \DOMXPath
    {
        $dom = new \DOMDocument();
        self::assertTrue($dom->loadXML($xml));
        $xpath = new \DOMXPath($dom);
        $xpath->registerNamespace('s', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
        return $xpath;
    }

    public function testRulesAttributesStylesAndLinksSurviveReverseSheetOrder(): void
    {
        $excel = Excel::template($this->tpl, $this->out);
        foreach (['Beta', 'Alpha'] as $name) {
            $excel->sheet($name)->fill(['{{NAME}}' => $name])->transferRows();
        }
        $excel->save();
        for ($i = 1; $i <= 2; $i++) {
            $xpath = $this->xpath($this->part($this->out, 'xl/worksheets/sheet' . $i . '.xml'));
            self::assertSame(1, $xpath->query('//s:conditionalFormatting/s:cfRule[@dxfId="0"]')->length);
            self::assertSame('Invalid & value', $xpath->evaluate('string(//s:dataValidation/@errorTitle)'));
            self::assertSame('Choose 1–9', $xpath->evaluate('string(//s:dataValidation/@error)'));
            self::assertSame('true', $xpath->evaluate('string(//s:dataValidation/@allowBlank)'));
            self::assertSame('42', $xpath->evaluate('string(//s:dataValidations/@xWindow)'));
            self::assertSame('B1 B2', $xpath->evaluate('string(//s:dataValidation/@sqref)'));
            self::assertSame('Internal link', $xpath->evaluate('string(//s:hyperlink/@tooltip)'));
            $children = [];
            foreach ($xpath->query('/s:worksheet/*') as $node) { $children[] = $node->localName; }
            self::assertLessThan(array_search('dataValidations', $children, true), array_search('conditionalFormatting', $children, true));
            self::assertLessThan(array_search('hyperlinks', $children, true), array_search('dataValidations', $children, true));
        }
        $styles = $this->xpath($this->part($this->out, 'xl/styles.xml'));
        self::assertSame(1, $styles->query('//s:dxfs/s:dxf/s:font/s:b')->length);
        self::assertSame('0.000', $styles->evaluate('string(//s:dxfs/s:dxf/s:numFmt/@formatCode)'));
        self::assertSame([], $excel->metadataDiagnostics());
    }

    public function testStrictModeRejectsUnmappedRangesAndPreservesDestination(): void
    {
        file_put_contents($this->out, 'previous result');
        $excel = Excel::template($this->tpl, $this->out, ['strict_metadata' => true]);
        $excel->sheet()->insertRow(['A' => 'inserted'])->transferRows();
        try {
            $excel->save();
            self::fail('Expected unsupported range transformation');
        } catch (Exception $e) {
            self::assertStringContainsString('metadata', $e->getMessage());
        }
        self::assertSame('previous result', file_get_contents($this->out));
        self::assertNotEmpty($excel->metadataDiagnostics());
    }

    public function testDefaultModeReportsUnmappedRangesWithoutChangingThem(): void
    {
        $excel = Excel::template($this->tpl, $this->out);
        $excel->sheet()->insertRow(['A' => 'inserted'])->transferRows();
        $excel->save();
        self::assertNotEmpty($excel->metadataDiagnostics());
        self::assertStringContainsString('sqref="B1:B2"', $this->part($this->out, 'xl/worksheets/sheet1.xml'));
    }

    public function testUnknownWorksheetFeatureIsReportedOrRejected(): void
    {
        $zip = new \ZipArchive();
        self::assertTrue($zip->open($this->tpl));
        $xml = $zip->getFromName('xl/worksheets/sheet1.xml');
        $xml = str_replace('</worksheet>', '<futureFeature xmlns="urn:unsupported"/></worksheet>', $xml);
        self::assertTrue($zip->addFromString('xl/worksheets/sheet1.xml', $xml));
        self::assertTrue($zip->close());
        $permissive = Excel::template($this->tpl, $this->out);
        $permissive->sheet()->transferRows();
        $permissive->save();
        self::assertSame('unsupported-node', $permissive->metadataDiagnostics()[0]['code']);
        self::assertStringNotContainsString('futureFeature', $this->part($this->out, 'xl/worksheets/sheet1.xml'));
        $excel = Excel::template($this->tpl, $this->out, ['strict_metadata' => true]);
        $excel->sheet()->transferRows();
        $this->expectException(Exception::class);
        $this->expectExceptionMessage('metadata');
        $excel->save();
    }

    public function testEmptySheetStillPreservesValidationMetadata(): void
    {
        $zip = new \ZipArchive();
        self::assertTrue($zip->open($this->tpl));
        $xml = $zip->getFromName('xl/worksheets/sheet1.xml');
        $xml = preg_replace('/<sheetData>.*?<\/sheetData>/s', '<sheetData/>', $xml);
        self::assertTrue($zip->addFromString('xl/worksheets/sheet1.xml', $xml));
        self::assertTrue($zip->close());
        $excel = Excel::template($this->tpl, $this->out);
        $excel->sheet()->transferRows();
        $excel->save();
        self::assertStringContainsString('errorTitle="Invalid &amp; value"', $this->part($this->out, 'xl/worksheets/sheet1.xml'));
    }

    public function testNewValidationIsAddedAlongsideTemplateValidation(): void
    {
        $excel = Excel::template($this->tpl, $this->out);
        $excel->sheet()->transferRows();
        $excel->sheet()->sheetWriter->addDataValidation('A1', \avadim\FastExcelWriter\DataValidation\DataValidation::dropDown(['Yes', 'No']));
        $excel->save();
        $xpath = $this->xpath($this->part($this->out, 'xl/worksheets/sheet1.xml'));
        self::assertSame(2, $xpath->query('//s:dataValidation')->length);
        self::assertSame('2', $xpath->evaluate('string(//s:dataValidations/@count)'));
        self::assertSame('42', $xpath->evaluate('string(//s:dataValidations/@xWindow)'));
    }

    public function testHeaderFooterTextAndFalseFlagsArePreserved(): void
    {
        $zip = new \ZipArchive();
        self::assertTrue($zip->open($this->tpl));
        $xml = $zip->getFromName('xl/worksheets/sheet1.xml');
        $xml = preg_replace('/<headerFooter\b[^>]*(?:\/>|>.*?<\/headerFooter>)/s', '<headerFooter differentFirst="false" differentOddEven="false" alignWithMargins="0"><oddHeader>&amp;CWords &amp;amp; text</oddHeader></headerFooter>', $xml);
        self::assertTrue($zip->addFromString('xl/worksheets/sheet1.xml', $xml));
        $zip->close();
        $excel = Excel::template($this->tpl, $this->out);
        $excel->sheet()->transferRows();
        $excel->save();
        $xpath = $this->xpath($this->part($this->out, 'xl/worksheets/sheet1.xml'));
        self::assertSame('&CWords &amp; text', $xpath->evaluate('string(//s:oddHeader)'));
        self::assertSame('false', $xpath->evaluate('string(//s:headerFooter/@differentFirst)'));
        self::assertSame('0', $xpath->evaluate('string(//s:headerFooter/@alignWithMargins)'));
    }

    public function testStrictFailureAfterOneSheetCleansMetadataTemporaryFiles(): void
    {
        $directory = sys_get_temp_dir() . '/fxt-meta-cleanup-' . uniqid('', true);
        self::assertTrue(mkdir($directory));
        $out = $directory . '/existing.xlsx';
        file_put_contents($out, 'previous result');
        try {
            $excel = Excel::template($this->tpl, $out, ['strict_metadata' => true]);
            $excel->sheet('Alpha')->transferRows();
            $excel->sheet('Beta')->insertRow(['A' => 'insert'])->transferRows();
            try {
                $excel->save();
                self::fail('Expected strict metadata rejection');
            } catch (Exception $e) {
                self::assertStringContainsString('metadata', $e->getMessage());
            }
            self::assertSame('previous result', file_get_contents($out));
            self::assertSame(['existing.xlsx'], array_map('basename', glob($directory . '/*')));
        } finally { @unlink($out); rmdir($directory); }
    }
}
