<?php

declare(strict_types=1);

namespace avadim\FastExcelTemplator;

use avadim\FastExcelWriter\Excel as Writer;
use PHPUnit\Framework\TestCase;

final class TemplateAssetsTest extends TestCase
{
    public function testImagesNotesTablesExtensionsAndRelationshipsSurvive(): void
    {
        $base = sys_get_temp_dir() . '/fxt-assets-' . uniqid('', true);
        $tpl = $base . '.xlsx';
        $out = $base . '-out.xlsx';
        $png = $base . '.png';
        file_put_contents($png, base64_decode('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII='));
        try {
            $writer = Writer::create(['Alpha', 'Beta']);
            foreach (['Alpha', 'Beta'] as $name) {
                $sheet = $writer->sheet($name);
                $sheet->writeRow(['Header', 'Quantity']);
                $sheet->writeRow(['{{VALUE}}', 2]);
                $sheet->addImage('D1', $png);
                $sheet->addNote('A2', 'Template note ' . $name);
            }
            $writer->save($tpl);
            $source = new \ZipArchive();
            self::assertTrue($source->open($tpl));
            $sheetXml = $source->getFromName('xl/worksheets/sheet1.xml');
            $sheetXml = str_replace('<worksheet ', '<worksheet xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" ', $sheetXml);
            $sheetXml = str_replace('</sheetData>', '</sheetData><hyperlinks><hyperlink ref="A2" r:id="rIdExternal" tooltip="External target"/></hyperlinks>', $sheetXml);
            $sheetXml = str_replace('</worksheet>', '<tableParts count="1"><tablePart r:id="rIdTable"/></tableParts>'
                . '<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"><x14:dataValidations xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main" count="1"><x14:dataValidation type="list" allowBlank="1"><x14:formula1><xm:f>Beta!$A$1:$A$2</xm:f></x14:formula1><xm:sqref>B2</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst></worksheet>', $sheetXml);
            self::assertTrue($source->addFromString('xl/worksheets/sheet1.xml', $sheetXml));
            $relationships = $source->getFromName('xl/worksheets/_rels/sheet1.xml.rels');
            $relationships = str_replace('</Relationships>', '<Relationship Id="rIdTable" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>', $relationships);
            $relationships = str_replace('</Relationships>', '<Relationship Id="rIdExternal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/" TargetMode="External"/></Relationships>', $relationships);
            self::assertTrue($source->addFromString('xl/worksheets/_rels/sheet1.xml.rels', $relationships));
            self::assertTrue($source->addFromString('xl/tables/table1.xml', '<?xml version="1.0"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Items" displayName="Items" ref="A1:B2" totalsRowShown="0"><autoFilter ref="A1:B2"/><tableColumns count="2"><tableColumn id="1" name="Header"/><tableColumn id="2" name="Quantity"/></tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>'));
            $types = $source->getFromName('[Content_Types].xml');
            self::assertTrue($source->addFromString('[Content_Types].xml', str_replace('</Types>', '<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>', $types)));
            $source->close();

            $excel = Excel::template($tpl, $out, ['strict_metadata' => true]);
            foreach (['Beta', 'Alpha'] as $name) {
                $excel->sheet($name)->fill(['{{VALUE}}' => $name])->transferRows();
            }
            $excel->save();
            self::assertSame([], $excel->metadataDiagnostics());

            self::assertTrue($source->open($tpl));
            $result = new \ZipArchive();
            self::assertTrue($result->open($out));
            $assets = 0;
            for ($i = 0; $i < $source->numFiles; $i++) {
                $entry = $source->getNameIndex($i);
                if (preg_match('#^xl/(media/|drawings/|comments|tables/|worksheets/_rels/)#', $entry)) {
                    self::assertSame($source->getFromName($entry), $result->getFromName($entry), $entry);
                    $assets++;
                }
            }
            self::assertGreaterThan(5, $assets);
            foreach ([1, 2] as $i) {
                $dom = new \DOMDocument();
                self::assertTrue($dom->loadXML($result->getFromName('xl/worksheets/sheet' . $i . '.xml')));
                $xpath = new \DOMXPath($dom);
                $xpath->registerNamespace('s', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
                $xpath->registerNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
                $rels = new \DOMDocument();
                self::assertTrue($rels->loadXML($result->getFromName('xl/worksheets/_rels/sheet' . $i . '.xml.rels')));
                $ids = [];
                foreach ($rels->documentElement->childNodes as $rel) {
                    if ($rel instanceof \DOMElement) { $ids[] = $rel->getAttribute('Id'); }
                }
                foreach ($xpath->query('//@r:id') as $id) { self::assertContains($id->value, $ids); }
                self::assertSame(1, $xpath->query('//s:drawing')->length);
                self::assertSame(1, $xpath->query('//s:legacyDrawing')->length);
                if ($i === 1) {
                    self::assertSame(1, $xpath->query('//s:tablePart')->length);
                    self::assertSame(1, $dom->getElementsByTagNameNS('http://schemas.microsoft.com/office/spreadsheetml/2009/9/main', 'dataValidations')->length);
                }
            }
            $source->close();
            $result->close();
        } finally {
            gc_collect_cycles();
            @unlink($png);
            @unlink($tpl);
            @unlink($out);
        }
    }
}
