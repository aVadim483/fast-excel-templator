<?php

declare(strict_types=1);
namespace avadim\FastExcelTemplator;

use avadim\FastExcelWriter\Excel as Writer;
use PHPUnit\Framework\TestCase;

final class CellMutationTest extends TestCase
{
    private string $tpl;
    private string $out;

    protected function setUp(): void
    {
        $this->tpl = sys_get_temp_dir() . '/fxt-cell-' . uniqid() . '.xlsx';
        $this->out = $this->tpl . '-out.xlsx';
    }

    protected function tearDown(): void
    {
        @unlink($this->tpl);
        @unlink($this->out);
    }

    public function testFormulaCanBeReplacedByScalarValues(): void
    {
        $values = [42, 0, 'text', '', true, false];
        $writer = Writer::create(['Sheet1']);
        foreach ($values as $value) {
            $writer->sheet()->writeRow(['=1+1'], null, ['A' => ['font-style' => 'bold']]);
        }
        $writer->save($this->tpl);
        $excel = Excel::template($this->tpl, $this->out);
        $excel->sheet()->rows(function ($src, $dst, $row) use ($values) {
            return $row->setValue('A', $values[$src - 1]);
        });
        $excel->save();
        $zip = new \ZipArchive();
        self::assertTrue($zip->open($this->out));
        $xml = $zip->getFromName('xl/worksheets/sheet1.xml');
        $zip->close();
        self::assertStringNotContainsString('<f', $xml);
        $cells = \avadim\FastExcelReader\Excel::open($this->out)->readCellsWithStyles();
        foreach ($values as $index => $value) {
            $cell = $cells['A' . ($index + 1)];
            self::assertSame($value, $cell['v']);
            self::assertNotEmpty($cell['s']['font']['font-style-bold']);
        }
    }

    public function testRemovingCellsMovesStyleAndMergeWithValue(): void
    {
        $writer = Writer::create(['Sheet1']);
        $sheet = $writer->sheet();
        $sheet->writeRow(['a', 'remove', 'c', 'd', 'remove', 12.5], null,
            ['C' => ['font-style' => 'bold'], 'F' => ['format' => '0.00']]);
        $sheet->mergeCells('C1:D1');
        $writer->save($this->tpl);
        $excel = Excel::template($this->tpl, $this->out);
        $excel->sheet()->rows(function ($src, $dst, $row) { return $row->removeCells(['B', 'E']); });
        $excel->save();
        $reader = \avadim\FastExcelReader\Excel::open($this->out);
        $cells = $reader->readCellsWithStyles();
        self::assertSame('c', $cells['B1']['v']);
        self::assertNotEmpty($cells['B1']['s']['font']['font-style-bold']);
        self::assertEmpty($cells['C1']['s']['font']['font-style-bold'] ?? null);
        self::assertSame(12.5, $cells['D1']['v']);
        self::assertSame('0.00', $cells['D1']['s']['format']['format-pattern']);
        self::assertSame('B1:C1', $reader->sheet()->mergedRange('B1'));
        self::assertArrayNotHasKey('F1', $cells);
    }

    public function testNullClearsFormulaButKeepsStyle(): void
    {
        $writer = Writer::create(['Sheet1']);
        $writer->sheet()->writeRow(['=1+1'], null, ['A' => ['font-style' => 'bold']]);
        $writer->save($this->tpl);
        $excel = Excel::template($this->tpl, $this->out);
        $excel->sheet()->rows(function ($src, $dst, $row) { return $row->setValue('A', null); });
        $excel->save();
        $cells = \avadim\FastExcelReader\Excel::open($this->out)->readCellsWithStyles();
        self::assertEmpty($cells['A1']['v']);
        self::assertNotEmpty($cells['A1']['s']['font']['font-style-bold']);
        $zip = new \ZipArchive();
        self::assertTrue($zip->open($this->out));
        $dom = new \DOMDocument();
        self::assertTrue($dom->loadXML($zip->getFromName('xl/worksheets/sheet1.xml')));
        $zip->close();
        self::assertSame(0, $dom->getElementsByTagName('f')->length);
        self::assertSame(0, $dom->getElementsByTagName('v')->length);
    }

    public function testScalarCanBecomeFormulaWithoutLosingStyle(): void
    {
        $writer = Writer::create(['Sheet1']);
        $writer->sheet()->writeRow([42], null, ['A' => ['font-style' => 'bold']]);
        $writer->save($this->tpl);
        $excel = Excel::template($this->tpl, $this->out);
        $excel->sheet()->rows(function ($src, $dst, $row) { return $row->setValue('A', '=2+3'); });
        $excel->save();
        $zip = new \ZipArchive();
        self::assertTrue($zip->open($this->out));
        self::assertStringContainsString('<f>2+3</f>', $zip->getFromName('xl/worksheets/sheet1.xml'));
        $zip->close();
        $cells = \avadim\FastExcelReader\Excel::open($this->out)->readCellsWithStyles();
        self::assertNotEmpty($cells['A1']['s']['font']['font-style-bold']);
    }

    public function testDeletionInsideHorizontalMergeShrinksItsRange(): void
    {
        $writer = Writer::create(['Sheet1']);
        $writer->sheet()->writeRow(['keep', 'anchor', '', '', 12.5]);
        $writer->sheet()->mergeCells('B1:D1');
        $writer->save($this->tpl);
        $excel = Excel::template($this->tpl, $this->out);
        $excel->sheet()->rows(function ($src, $dst, $row) { return $row->removeCell('C'); });
        $excel->save();
        $reader = \avadim\FastExcelReader\Excel::open($this->out);
        self::assertSame('B1:C1', $reader->sheet()->mergedRange('B1'));
        self::assertSame(12.5, $reader->readRows()[1]['D']);
    }

    public function testSparseFormulaMovesWithRelativeReferencesAndNumberFormat(): void
    {
        $writer = Writer::create(['Sheet1']);
        $writer->sheet()->writeRow(['left', 10, '', '', '=B1'], null, ['E' => ['format' => '0.00']]);
        $writer->save($this->tpl);
        $zip = new \ZipArchive();
        self::assertTrue($zip->open($this->tpl));
        $xml = preg_replace('/<c\b[^>]*\br="[CD]1"[^>]*>.*?<\/c>/s', '', $zip->getFromName('xl/worksheets/sheet1.xml'));
        self::assertTrue($zip->addFromString('xl/worksheets/sheet1.xml', $xml));
        $zip->close();
        $excel = Excel::template($this->tpl, $this->out);
        $excel->sheet()->rows(function ($src, $dst, $row) { return $row->removeCell('B'); });
        $excel->save();
        self::assertTrue($zip->open($this->out));
        self::assertStringContainsString('<f>A1</f>', $zip->getFromName('xl/worksheets/sheet1.xml'));
        $zip->close();
        $cells = \avadim\FastExcelReader\Excel::open($this->out)->readCellsWithStyles();
        self::assertSame('0.00', $cells['D1']['s']['format']['format-pattern']);
    }

    public function testDeletionAffectingMultiRowMergeIsExplicitlyRejected(): void
    {
        $writer = Writer::create(['Sheet1']);
        $writer->sheet()->writeRow(['remove', 'anchor', '']);
        $writer->sheet()->writeRow(['remove', '', '']);
        $writer->sheet()->mergeCells('B1:C2');
        $writer->save($this->tpl);
        $excel = Excel::template($this->tpl, $this->out);
        $this->expectException(Exception::class);
        $this->expectExceptionMessage('multi-row merged range');
        $excel->sheet()->rows(function ($src, $dst, $row) { return $row->removeCell('A'); });
    }
}
