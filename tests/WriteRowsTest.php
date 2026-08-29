<?php

declare(strict_types=1);

namespace avadim\FastExcelTemplator;

use avadim\FastExcelWriter\Excel as Writer;
use PHPUnit\Framework\TestCase;

/**
 * The row-level write API of SheetTemplate that had no coverage yet:
 * replaceRow() and cloneRow() (both work against the source cursor), writeRow() and
 * writeCell() (both write past the template, straight to the output), and the cell
 * removal of RowTemplate.
 */
final class WriteRowsTest extends TestCase
{
    private string $tpl;
    private string $out;

    protected function setUp(): void
    {
        $name = preg_replace('/\W/', '', $this->getName());
        $dir = sys_get_temp_dir() . DIRECTORY_SEPARATOR;
        $this->tpl = $dir . 'fxt-write-' . $name . '-tpl.xlsx';
        $this->out = $dir . 'fxt-write-' . $name . '-out.xlsx';
        @unlink($this->tpl);
        @unlink($this->out);
    }

    protected function tearDown(): void
    {
        @unlink($this->tpl);
        @unlink($this->out);
    }

    /**
     * A template of $countRows rows, three columns wide: 'h1..h3', then 'r2a..r2c', 'r3a..r3c', ...
     */
    private function makeTemplate(int $countRows = 3): void
    {
        $writer = Writer::create(['Sheet1']);
        $sheet = $writer->sheet();
        $sheet->writeRow(['h1', 'h2', 'h3']);
        for ($rowNum = 2; $rowNum <= $countRows; $rowNum++) {
            $sheet->writeRow(['r' . $rowNum . 'a', 'r' . $rowNum . 'b', 'r' . $rowNum . 'c']);
        }
        $writer->save($this->tpl);
    }

    private function outputRows(): array
    {
        self::assertTrue(is_file($this->out));

        return \avadim\FastExcelReader\Excel::open($this->out)->readRows();
    }

    /**
     * replaceRow() writes its own row and consumes one source row, so the replacement
     * takes the place of the template row instead of being added next to it
     */
    public function testReplaceRowTakesThePlaceOfTheSourceRow()
    {
        $this->makeTemplate();

        $excel = Excel::template($this->tpl, $this->out);
        $sheet = $excel->sheet();
        $sheet->transferRows(1);
        $sheet->replaceRow(['A' => 'X', 'B' => 'Y']);
        $sheet->transferRows();
        $excel->save();

        $rows = $this->outputRows();

        // the header is untouched, row 2 is the replacement, row 3 is the source row 3
        self::assertSame(['A' => 'h1', 'B' => 'h2', 'C' => 'h3'], $rows[1]);
        self::assertSame(['A' => 'X', 'B' => 'Y'], $rows[2]);
        self::assertSame(['A' => 'r3a', 'B' => 'r3b', 'C' => 'r3c'], $rows[3]);
        // the sheet keeps as many rows as the template had
        self::assertCount(3, $rows);
    }

    public function testReplaceRowIsFluent()
    {
        $this->makeTemplate();

        $excel = Excel::template($this->tpl, $this->out);
        $sheet = $excel->sheet();

        self::assertSame($sheet, $sheet->replaceRow(['A' => 'X']));
    }

    /**
     * cloneRow() re-inserts the row that was read last, so the source row appears twice —
     * once as transferred, once as the clone with its own data
     */
    public function testCloneRowRepeatsTheLastReadRow()
    {
        $this->makeTemplate();

        $excel = Excel::template($this->tpl, $this->out);
        $sheet = $excel->sheet();
        $sheet->transferRows(2);
        $sheet->cloneRow(['B' => 'cloned']);
        $sheet->transferRows();
        $excel->save();

        $rows = $this->outputRows();

        self::assertCount(4, $rows);
        self::assertSame(['A' => 'r2a', 'B' => 'r2b', 'C' => 'r2c'], $rows[2]);
        // the clone keeps the cells of row 2 and takes the new value of 'B'
        self::assertSame(['A' => 'r2a', 'B' => 'cloned', 'C' => 'r2c'], $rows[3]);
        self::assertSame(['A' => 'r3a', 'B' => 'r3b', 'C' => 'r3c'], $rows[4]);
    }

    public function testCloneRowIsFluent()
    {
        $this->makeTemplate();

        $excel = Excel::template($this->tpl, $this->out);
        $sheet = $excel->sheet();
        $sheet->transferRows(1);

        self::assertSame($sheet, $sheet->cloneRow());
    }

    /**
     * writeRow() and writeCell() do not touch the source cursor: they write at the current
     * output position, so they can add rows the template does not have
     */
    public function testWriteRowAndWriteCellAppendToTheOutput()
    {
        $this->makeTemplate(1);

        $excel = Excel::template($this->tpl, $this->out);
        $sheet = $excel->sheet();
        $sheet->transferRows(1);

        self::assertSame($sheet, $sheet->writeRow(['w1', 'w2', 'w3']));
        self::assertSame($sheet, $sheet->writeCell('c1'));
        self::assertSame($sheet, $sheet->writeCell('c2'));

        $excel->save();

        $rows = $this->outputRows();

        self::assertSame(['A' => 'h1', 'B' => 'h2', 'C' => 'h3'], $rows[1]);
        self::assertSame(['A' => 'w1', 'B' => 'w2', 'C' => 'w3'], $rows[2]);
        // writeCell() moves the pointer to the next cell of the same row
        self::assertSame(['A' => 'c1', 'B' => 'c2'], $rows[3]);
    }

    public function testWriteRowAppliesStyles()
    {
        $this->makeTemplate(1);

        $excel = Excel::template($this->tpl, $this->out);
        $sheet = $excel->sheet();
        $sheet->transferRows(1);
        $sheet->writeRow(['bold', 'plain'], null, ['A' => ['font-style' => 'bold']]);
        $excel->save();

        $cells = \avadim\FastExcelReader\Excel::open($this->out)->readCellsWithStyles();

        self::assertSame('bold', $cells['A2']['v']);
        self::assertNotEmpty($cells['A2']['s']['font']['font-style-bold'] ?? null);
        self::assertArrayNotHasKey('font-style-bold', $cells['B2']['s']['font'] ?? []);
    }

    /**
     * A removed cell is not written and the cells to its right move one column left
     */
    public function testRemoveCellShiftsTheRestOfTheRowLeft()
    {
        $this->makeTemplate(2);

        $excel = Excel::template($this->tpl, $this->out);
        $excel->sheet()->rows(function ($sourceRowNum, $targetRowNum, RowTemplate $row) {
            $row->removeCell('B');

            return $row;
        });
        $excel->save();

        $rows = $this->outputRows();

        self::assertSame(['A' => 'h1', 'B' => 'h3'], $rows[1]);
        self::assertSame(['A' => 'r2a', 'B' => 'r2c'], $rows[2]);
    }

    /**
     * removeCells() removes several columns at once and, like the rest of RowTemplate,
     * accepts column letters in any case
     */
    public function testRemoveCellsRemovesEveryListedColumn()
    {
        $this->makeTemplate(2);

        $excel = Excel::template($this->tpl, $this->out);
        $excel->sheet()->rows(function ($sourceRowNum, $targetRowNum, RowTemplate $row) {
            $row->removeCells(['b', 'C']);

            return $row;
        });
        $excel->save();

        $rows = $this->outputRows();

        self::assertSame(['A' => 'h1'], $rows[1]);
        self::assertSame(['A' => 'r2a'], $rows[2]);
    }

    /**
     * Removing a column the row does not have is a no-op, not an error
     */
    public function testRemoveCellsIgnoresUnknownColumns()
    {
        $row = new RowTemplate();
        $row->addCell('A', ['v' => 'kept']);

        self::assertSame($row, $row->removeCells(['Z', 'AA']));
        self::assertSame(['A'], array_keys($row->cells()));
        self::assertArrayNotHasKey('__removed', $row->cells()['A']);
    }
}
