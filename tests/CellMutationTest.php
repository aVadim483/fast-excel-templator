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
}
