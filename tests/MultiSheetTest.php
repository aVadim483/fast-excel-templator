<?php

declare(strict_types=1);

namespace avadim\FastExcelTemplator;

use avadim\FastExcelWriter\Excel as Writer;
use PHPUnit\Framework\TestCase;

/**
 * A workbook of several sheets: every sheet keeps its own read cursor and its own writer,
 * so they can be filled independently and in any order, and save() must splice all of them
 * back into the template.
 */
final class MultiSheetTest extends TestCase
{
    private string $tpl;
    private string $out;

    protected function setUp(): void
    {
        $name = preg_replace('/\W/', '', $this->getName());
        $dir = sys_get_temp_dir() . DIRECTORY_SEPARATOR;
        $this->tpl = $dir . 'fxt-multi-' . $name . '-tpl.xlsx';
        $this->out = $dir . 'fxt-multi-' . $name . '-out.xlsx';
        @unlink($this->tpl);
        @unlink($this->out);
    }

    protected function tearDown(): void
    {
        @unlink($this->tpl);
        @unlink($this->out);
    }

    /**
     * Two sheets, 'Alpha' and 'Beta', each with a header and two data rows
     */
    private function makeTemplate(): void
    {
        $writer = Writer::create(['Alpha', 'Beta']);

        $alpha = $writer->sheet('Alpha');
        $alpha->writeRow(['{{TITLE}}', 'alpha']);
        $alpha->writeRow(['a1', 'a2']);
        $alpha->writeRow(['a3', 'a4']);

        $beta = $writer->sheet('Beta');
        $beta->writeRow(['{{TITLE}}', 'beta']);
        $beta->writeRow(['b1', 'b2']);
        $beta->writeRow(['b3', 'b4']);

        $writer->save($this->tpl);
    }

    private function outputRows(string $sheetName): array
    {
        self::assertTrue(is_file($this->out));

        return \avadim\FastExcelReader\Excel::open($this->out)->sheet($sheetName)->readRows();
    }

    public function testEverySheetOfTheTemplateIsImported()
    {
        $this->makeTemplate();

        $excel = Excel::template($this->tpl, $this->out);
        $sheets = $excel->sheets();

        self::assertCount(2, $sheets);
        self::assertSame(['Alpha', 'Beta'], array_values(array_map(function (SheetTemplate $sheet) {
            return $sheet->name();
        }, $sheets)));
        self::assertSame('Alpha', $excel->sheet()->name());
        self::assertSame('Beta', $excel->sheet('Beta')->name());
    }

    public function testEverySheetIsTransferredToTheOutput()
    {
        $this->makeTemplate();

        $excel = Excel::template($this->tpl, $this->out);
        foreach ($excel->sheets() as $sheet) {
            $sheet->transferRows();
        }
        $excel->save();

        self::assertSame(
            [
                1 => ['A' => '{{TITLE}}', 'B' => 'alpha'],
                2 => ['A' => 'a1', 'B' => 'a2'],
                3 => ['A' => 'a3', 'B' => 'a4'],
            ],
            $this->outputRows('Alpha')
        );
        self::assertSame(
            [
                1 => ['A' => '{{TITLE}}', 'B' => 'beta'],
                2 => ['A' => 'b1', 'B' => 'b2'],
                3 => ['A' => 'b3', 'B' => 'b4'],
            ],
            $this->outputRows('Beta')
        );
    }

    /**
     * The sheets do not share a cursor: rows inserted into one of them must not shift
     * or consume anything in the other, and the order in which they are filled is free
     */
    public function testSheetsAreFilledIndependently()
    {
        $this->makeTemplate();

        $excel = Excel::template($this->tpl, $this->out);

        // the second sheet is processed first, and gets an extra row
        $beta = $excel->sheet('Beta');
        $beta->transferRows(1);
        $beta->insertRow(['A' => 'inserted', 'B' => 'into beta']);
        $beta->transferRows();

        // the first one is transferred as is, afterwards
        $excel->sheet('Alpha')->transferRows();

        $excel->save();

        self::assertSame(
            [
                1 => ['A' => '{{TITLE}}', 'B' => 'alpha'],
                2 => ['A' => 'a1', 'B' => 'a2'],
                3 => ['A' => 'a3', 'B' => 'a4'],
            ],
            $this->outputRows('Alpha')
        );
        self::assertSame(
            [
                1 => ['A' => '{{TITLE}}', 'B' => 'beta'],
                2 => ['A' => 'inserted', 'B' => 'into beta'],
                3 => ['A' => 'b1', 'B' => 'b2'],
                4 => ['A' => 'b3', 'B' => 'b4'],
            ],
            $this->outputRows('Beta')
        );
    }

    /**
     * Excel::fill() is the workbook-wide shortcut: it registers the replacement on every sheet
     */
    public function testFillOnTheWorkbookReachesEverySheet()
    {
        $this->makeTemplate();

        $excel = Excel::template($this->tpl, $this->out);
        $excel->fill(['{{TITLE}}' => 'Report 2026']);
        foreach ($excel->sheets() as $sheet) {
            $sheet->transferRows();
        }
        $excel->save();

        self::assertSame('Report 2026', $this->outputRows('Alpha')[1]['A']);
        self::assertSame('Report 2026', $this->outputRows('Beta')[1]['A']);
    }

    /**
     * A sheet left untouched still has to reach the output — as the empty sheet it is,
     * not as a missing part of the workbook
     */
    public function testUntouchedSheetIsStillPresentInTheOutput()
    {
        $this->makeTemplate();

        $excel = Excel::template($this->tpl, $this->out);
        $excel->sheet('Alpha')->transferRows();
        $excel->save();

        $reader = \avadim\FastExcelReader\Excel::open($this->out);

        self::assertSame(['Alpha', 'Beta'], array_values($reader->getSheetNames()));
        self::assertCount(3, $reader->sheet('Alpha')->readRows());
        self::assertSame([], $reader->sheet('Beta')->readRows());
    }

    public function testWorkbookOrderCanDifferFromWorksheetPartNumbers(): void
    {
        $this->makeTemplate();
        $zip = new \ZipArchive();
        self::assertTrue($zip->open($this->tpl));
        $dom = new \DOMDocument();
        self::assertTrue($dom->loadXML($zip->getFromName('xl/workbook.xml')));
        $sheets = $dom->getElementsByTagName('sheets')->item(0);
        $first = $sheets->getElementsByTagName('sheet')->item(0);
        $sheets->appendChild($first);
        self::assertTrue($zip->addFromString('xl/workbook.xml', $dom->saveXML()));
        $zip->close();
        $excel = Excel::template($this->tpl, $this->out);
        $excel->sheet('Beta')->transferRows();
        $excel->sheet('Alpha')->transferRows();
        $excel->save();
        self::assertSame('a1', $this->outputRows('Alpha')[2]['A']);
        self::assertSame('b1', $this->outputRows('Beta')[2]['A']);
    }
}
