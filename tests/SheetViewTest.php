<?php

declare(strict_types=1);

namespace avadim\FastExcelTemplator;

use avadim\FastExcelWriter\Excel as Writer;
use PHPUnit\Framework\TestCase;

/**
 * The <sheetViews> section of the template (frozen panes, active cell, grid lines) is captured
 * at open time and written back on save. It used to be reshaped inside getSheetViews(), i.e. a
 * getter changed the object it was asked about; the reshaping now happens where the data arrives.
 */
final class SheetViewTest extends TestCase
{
    private string $tpl;
    private string $out;

    protected function setUp(): void
    {
        $name = preg_replace('/\W/', '', $this->getName());
        $dir = sys_get_temp_dir() . DIRECTORY_SEPARATOR;
        $this->tpl = $dir . 'fxt-view-' . $name . '-tpl.xlsx';
        $this->out = $dir . 'fxt-view-' . $name . '-out.xlsx';
        @unlink($this->tpl);
        @unlink($this->out);
    }

    protected function tearDown(): void
    {
        @unlink($this->tpl);
        @unlink($this->out);
    }

    private function makeTemplate(): void
    {
        $writer = Writer::create(['Sheet1']);
        $sheet = $writer->sheet();
        $sheet->writeRow(['h1', 'h2']);
        $sheet->writeRow([1, 2]);
        $sheet->setFreezeRows(1);
        $writer->save($this->tpl);
    }

    private function sheetViews(string $file): string
    {
        $zip = new \ZipArchive();
        self::assertTrue($zip->open($file) === true);
        $xml = (string)$zip->getFromName('xl/worksheets/sheet1.xml');
        $zip->close();

        preg_match('#<sheetViews>.*?</sheetViews>#s', $xml, $m);

        return $m[0] ?? '';
    }

    public function testFrozenPaneIsTransferred()
    {
        $this->makeTemplate();

        $excel = Excel::template($this->tpl, $this->out);
        $excel->sheet()->transferRows();
        $excel->save();

        $out = $this->sheetViews($this->out);

        self::assertStringContainsString('state="frozen"', $out);
        self::assertStringContainsString('ySplit="1"', $out);
        // the whole section survives the round trip unchanged
        self::assertSame($this->sheetViews($this->tpl), $out);
    }

    /**
     * Asking twice must give the same answer: the getter must not consume what it returns
     */
    public function testGetSheetViewsIsRepeatable()
    {
        $this->makeTemplate();

        $sheet = Excel::template($this->tpl, $this->out)->sheet();

        $first = $sheet->sheetWriter->getSheetViews();
        $second = $sheet->sheetWriter->getSheetViews();

        self::assertNotEmpty($first);
        self::assertSame($first, $second);
    }

    /**
     * ...and it must not rewrite the object it is asked about. The old getter unwrapped the
     * stored structure in place, which stayed unnoticed only because doing it twice was harmless.
     */
    public function testGetSheetViewsLeavesTheObjectUnchanged()
    {
        $this->makeTemplate();

        $sheetWriter = Excel::template($this->tpl, $this->out)->sheet()->sheetWriter;

        $property = new \ReflectionProperty(\avadim\FastExcelWriter\Sheet::class, 'sheetViews');
        $property->setAccessible(true);

        $before = $property->getValue($sheetWriter);
        $sheetWriter->getSheetViews();
        $after = $property->getValue($sheetWriter);

        self::assertSame($before, $after, 'getSheetViews() must not modify $sheetViews');
    }
}
