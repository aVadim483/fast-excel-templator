<?php

declare(strict_types=1);

namespace avadim\FastExcelTemplator;

use avadim\FastExcelWriter\Excel as Writer;
use PHPUnit\Framework\TestCase;

/**
 * Merged cells travel with an inserted row: a merge captured in a row template has to be
 * recreated at the row the template lands on. Internally that is kept as relative offsets
 * (SheetTemplate::mergedRangeOffsets()), while the public mergedRange() must keep answering
 * exactly like the reader it is inherited from.
 */
final class MergedCellsTest extends TestCase
{
    private string $tpl;
    private string $out;

    protected function setUp(): void
    {
        // a template stays open while the Excel object lives, and an open file
        // cannot be replaced on Windows, hence a separate name per test
        $name = preg_replace('/\W/', '', $this->getName());
        $dir = sys_get_temp_dir() . DIRECTORY_SEPARATOR;
        $this->tpl = $dir . 'fxt-merged-' . $name . '-tpl.xlsx';
        $this->out = $dir . 'fxt-merged-' . $name . '-out.xlsx';
        @unlink($this->tpl);
        @unlink($this->out);
    }

    protected function tearDown(): void
    {
        @unlink($this->tpl);
        @unlink($this->out);
    }

    /**
     * Template: row 1 is a header, row 2 holds a merge of A2:C2
     */
    private function makeTemplate(): void
    {
        $writer = Writer::create(['Sheet1']);
        $sheet = $writer->sheet();
        $sheet->writeRow(['Header', '', '', 'Note']);
        $sheet->writeRow(['merged', '', '', 'tail']);
        $sheet->mergeCells('A2:C2');
        $writer->save($this->tpl);
    }

    private function mergeRefs(string $file): array
    {
        $zip = new \ZipArchive();
        self::assertTrue($zip->open($file) === true);
        $xml = (string)$zip->getFromName('xl/worksheets/sheet1.xml');
        $zip->close();

        preg_match_all('#<mergeCell ref="([^"]+)"#', $xml, $m);

        return $m[1];
    }

    /**
     * The inherited method must answer for EVERY cell of a merge and return an address range —
     * the templator used to override it with relative offsets and a null for non-corner cells.
     */
    public function testMergedRangeBehavesLikeInTheReader()
    {
        $this->makeTemplate();

        $sheet = Excel::template($this->tpl, $this->out)->sheet();

        self::assertSame('A2:C2', $sheet->mergedRange('A2'), 'top-left cell of the merge');
        self::assertSame('A2:C2', $sheet->mergedRange('B2'), 'cell inside the merge');
        self::assertSame('A2:C2', $sheet->mergedRange('C2'), 'last cell of the merge');
        self::assertNull($sheet->mergedRange('D2'), 'cell outside any merge');

        // and it answers the same as the plain reader does for the same file
        $reader = \avadim\FastExcelReader\Excel::open($this->tpl)->sheet();
        foreach (['A2', 'B2', 'C2', 'D2'] as $address) {
            self::assertSame($reader->mergedRange($address), $sheet->mergedRange($address), $address);
        }
    }

    /**
     * A merge of the template row is recreated on every row the template is inserted into.
     */
    public function testMergeFollowsInsertedRows()
    {
        $this->makeTemplate();

        $excel = Excel::template($this->tpl, $this->out);
        $sheet = $excel->sheet();
        $sheet->transferRowsUntil(1);
        $rowTemplate = $sheet->getRowTemplate(2);
        for ($i = 0; $i < 3; $i++) {
            $sheet->insertRow($rowTemplate, ['A' => 'row ' . $i]);
        }
        $excel->save();

        self::assertSame(['A2:C2', 'A3:C3', 'A4:C4'], $this->mergeRefs($this->out));
    }

    /**
     * A merge is preserved when rows are transferred as they are.
     */
    public function testMergeSurvivesPlainTransfer()
    {
        $this->makeTemplate();

        $excel = Excel::template($this->tpl, $this->out);
        $excel->sheet()->transferRows();
        $excel->save();

        self::assertSame(['A2:C2'], $this->mergeRefs($this->out));
    }
}
