<?php

declare(strict_types=1);

namespace avadim\FastExcelTemplator;

use avadim\FastExcelWriter\Excel as Writer;
use PHPUnit\Framework\TestCase;

final class StreamingContractsTest extends TestCase
{
    private string $tpl;
    private string $out;

    protected function setUp(): void
    {
        $this->tpl = sys_get_temp_dir() . '/fxt-stream-' . uniqid('', true) . '.xlsx';
        $this->out = $this->tpl . '-out.xlsx';
        $writer = Writer::create(['Sheet1']);
        for ($i = 1; $i <= 4; $i++) {
            $writer->sheet()->writeRow(['row ' . $i, 'default ' . $i, '=A' . $i]);
        }
        $writer->save($this->tpl);
    }

    protected function tearDown(): void
    {
        @unlink($this->tpl);
        @unlink($this->out);
    }

    private function sheetXml(): string
    {
        $zip = new \ZipArchive();
        self::assertTrue($zip->open($this->out));
        $xml = $zip->getFromName('xl/worksheets/sheet1.xml');
        $zip->close();
        return $xml;
    }

    public function testSingleTemplateIsNotMutatedByInsert(): void
    {
        $excel = Excel::template($this->tpl, $this->out);
        $sheet = $excel->sheet();
        $row = (new RowTemplate())->addCell('A', 'original')->addCell('B', 'default');
        $sheet->insertRow($row, ['A' => 1, 'B' => 'first']);
        $sheet->insertRow($row, ['A' => 2]);
        $excel->save();
        $rows = \avadim\FastExcelReader\Excel::open($this->out)->readRows();
        self::assertSame('default', $rows[2]['B']);
        self::assertSame('original', $row->getValue('A'));
        self::assertSame('default', $row->getValue('B'));
    }

    public function testCollectionRetainsDefaultsAndFormulasAcrossCyclesAfterForeach(): void
    {
        $excel = Excel::template($this->tpl, $this->out);
        $sheet = $excel->sheet();
        $templates = $sheet->getRowTemplates(1, 2, true);
        foreach ($templates as $row) {
            self::assertNotNull($row->getValue('B'));
        }
        $sheet->insertRow($templates, ['A' => 10, 'B' => 'override', 'C' => 42]);
        $sheet->insertRow($templates, ['A' => 20]);
        $sheet->insertRow($templates, ['A' => 30]);
        $sheet->insertRow($templates, ['A' => 40]);
        $excel->save();
        $rows = \avadim\FastExcelReader\Excel::open($this->out)->readRows();
        self::assertSame('default 1', $rows[3]['B']);
        self::assertSame('default 2', $rows[4]['B']);
        self::assertStringContainsString('<f>A3</f>', $this->sheetXml());
    }

    public function testZeroOperationsLeaveBothCursorsUntouched(): void
    {
        $excel = Excel::template($this->tpl, $this->out);
        $sheet = $excel->sheet();
        $sheet->transferRows(1);
        $written = $sheet->lastWrittenRowNum();
        $sheet->transferRows(0, function () { self::fail('Zero transfer invokes callback'); });
        $sheet->skipRows(0);
        self::assertSame(1, $sheet->lastReadRowNum());
        self::assertSame($written, $sheet->lastWrittenRowNum());
        $sheet->transferRows();
        $excel->save();
        self::assertSame('row 2', \avadim\FastExcelReader\Excel::open($this->out)->readRows()[2]['A']);
    }

    /** @dataProvider negativeOperations */
    public function testNegativeLimitsFailBeforeMovingCursor(string $method): void
    {
        $excel = Excel::template($this->tpl, $this->out);
        $sheet = $excel->sheet();
        try {
            $sheet->$method(-1);
            self::fail('Negative limit was accepted');
        } catch (Exception $e) {
            self::assertSame(0, $sheet->lastReadRowNum());
            self::assertSame(0, $sheet->lastWrittenRowNum());
        }
    }

    public static function negativeOperations(): array
    {
        return [['transferRows'], ['skipRows'], ['transferRowsUntil'], ['skipRowsUntil']];
    }

    private function makeSparseTemplate(bool $dimension): void
    {
        $zip = new \ZipArchive();
        self::assertTrue($zip->open($this->tpl));
        $xml = $zip->getFromName('xl/worksheets/sheet1.xml');
        $xml = preg_replace('/<row\b[^>]*\br="[23]"[^>]*>.*?<\/row>/s', '', $xml);
        if (!$dimension) {
            $xml = preg_replace('/<dimension\b[^>]*\/>/', '', $xml);
        }
        self::assertTrue($zip->addFromString('xl/worksheets/sheet1.xml', $xml));
        self::assertTrue($zip->close());
    }

    public function testSparseBoundaryDoesNotConsumeTheNextRow(): void
    {
        $this->makeSparseTemplate(true);
        $excel = Excel::template($this->tpl, $this->out);
        $sheet = $excel->sheet();
        $sheet->transferRowsUntil(2);
        self::assertSame(2, $sheet->lastReadRowNum());
        $sheet->skipRowsUntil(3);
        self::assertSame(3, $sheet->lastReadRowNum());
        $sheet->transferRows();
        $sheet->transferRows(); // EOF is repeatable
        $excel->save();
        self::assertSame('row 4', \avadim\FastExcelReader\Excel::open($this->out)->readRows()[4]['A']);
    }

    public function testMissingRowsInTemplateRangeAreEmptyTemplatesWithoutDimension(): void
    {
        $this->makeSparseTemplate(false);
        $excel = Excel::template($this->tpl, $this->out);
        $sheet = $excel->sheet();
        $templates = $sheet->getRowTemplates(1, 4, true);
        $values = [];
        foreach ($templates as $number => $row) {
            $values[$number] = $row->getValue('A');
        }
        self::assertSame([1 => 'row 1', 2 => null, 3 => null, 4 => 'row 4'], $values);
        self::assertSame(0, $sheet->lastReadRowNum());
    }

    public function testCallbackStopConsumesOnlyTheStoppedRow(): void
    {
        $excel = Excel::template($this->tpl, $this->out);
        $sheet = $excel->sheet();
        $sheet->transferRows(2, function ($src, $dst, $row) { return $src === 2 ? false : $row; });
        self::assertSame(2, $sheet->lastReadRowNum());
        $seen = [];
        $sheet->transferRows(null, function ($src, $dst, $row) use (&$seen) { $seen[] = $src; return $row; });
        self::assertSame([3, 4], $seen);
    }

    public function testSkipCountsAndUntilContinueAtTheNextSourcePosition(): void
    {
        $excel = Excel::template($this->tpl, $this->out);
        $sheet = $excel->sheet();
        $sheet->skipRows(1);
        $sheet->skipRowsUntil(2);
        $seen = [];
        $sheet->transferRows(1, function ($src, $dst, $row) use (&$seen) { $seen[] = $src; return $row; });
        self::assertSame([3], $seen);
        $sheet->skipRows();
        self::assertSame(4, $sheet->lastReadRowNum());
        $sheet->transferRows();
        self::assertSame([3], $seen);
    }

    public function testCallbackSkipAtBoundaryDoesNotReadBeyondLimit(): void
    {
        $excel = Excel::template($this->tpl, $this->out);
        $sheet = $excel->sheet();
        $seen = [];
        $sheet->transferRows(2, function ($src) use (&$seen) { $seen[] = $src; return null; });
        self::assertSame([1, 2], $seen);
        self::assertSame(2, $sheet->lastReadRowNum());
        $next = [];
        foreach ($sheet->readRow() as $src => $row) { $next[] = $src; }
        self::assertSame([3, 4], $next);
    }

    public function testEmptySheetAndRepeatedEndOfFileOperations(): void
    {
        $writer = Writer::create(['Sheet1']);
        $writer->save($this->tpl);
        $excel = Excel::template($this->tpl, $this->out);
        $sheet = $excel->sheet();
        $sheet->transferRows(0)->skipRows(0)->transferRows()->skipRows()->transferRows(5);
        self::assertSame(0, $sheet->lastReadRowNum());
        self::assertSame(0, $sheet->lastWrittenRowNum());
        self::assertTrue($excel->save());
    }

    public function testEmptyCollectionGivesAnExplicitException(): void
    {
        $excel = Excel::template($this->tpl, $this->out);
        $this->expectException(Exception::class);
        $this->expectExceptionMessage('empty template collection');
        $excel->sheet()->insertRow(new RowTemplateCollection());
    }

    public function testSelfClosingTemplateRowDoesNotStealTheFollowingRowsCells(): void
    {
        $zip = new \ZipArchive();
        self::assertTrue($zip->open($this->tpl));
        $xml = $zip->getFromName('xl/worksheets/sheet1.xml');
        $xml = preg_replace('/<row\b[^>]*\br="2"[^>]*>.*?<\/row>/s', '<row r="2" ht="25" customHeight="1"/>', $xml);
        self::assertTrue($zip->addFromString('xl/worksheets/sheet1.xml', $xml));
        self::assertTrue($zip->close());
        $excel = Excel::template($this->tpl, $this->out);
        $templates = $excel->sheet()->getRowTemplates(2, 3, true);
        $rows = iterator_to_array($templates);
        self::assertNull($rows[2]->getValue('A'));
        self::assertSame('25', $rows[2]->attribute('ht'));
        self::assertSame('row 3', $rows[3]->getValue('A'));
    }

    /** @dataProvider invalidRanges */
    public function testInvalidTemplateRangesFailWithoutMovingCursor(int $min, int $max): void
    {
        $excel = Excel::template($this->tpl, $this->out);
        $sheet = $excel->sheet();
        try {
            $sheet->getRowTemplates($min, $max);
            self::fail('Invalid range accepted');
        } catch (Exception $e) {
            self::assertSame(0, $sheet->lastReadRowNum());
        }
    }

    public static function invalidRanges(): array
    {
        return [[0, 1], [-1, 1], [3, 2], [1, 10]];
    }
}
