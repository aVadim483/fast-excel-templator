<?php

declare(strict_types=1);

namespace avadim\FastExcelTemplator;

use PHPUnit\Framework\TestCase;

final class RowTemplateTest extends TestCase
{
    const INP_FILE = __DIR__ . '/test_files/test.xlsx';
    const OUT_FILE = __DIR__ . '/test_files/test-out.xlsx';


    protected function template(): Excel
    {
        if (is_file(self::OUT_FILE)) {
            unlink(self::OUT_FILE);
        }

        return Excel::template(self::INP_FILE, self::OUT_FILE);
    }


    protected function read()
    {
        $this->assertTrue(is_file(self::OUT_FILE));

        $excelReader = \avadim\FastExcelReader\Excel::open(self::OUT_FILE);

        return $excelReader->readRowsWithStyles()[1];
    }

    public function testAddCloneAppendAndValues()
    {
        $excel = $this->template();
        $sheet = $excel->sheet();

        $sheet->rows(function ($sourceRowNum, $targetRowNum, RowTemplate $row) {
            // set directly to 'A'
            $row->setValue('A', 'foo');

            // lower-case should be normalized to 'B'
            $row->setValue('b', 123);

            // add cell 'C'
            $row->addCell('C', ['v' => 'X', 't' => 's']);

            // clone C to D and E
            $row->cloneCell('C', ['D', 'E']);

            // append the next cell after E
            // so withValue() assigns to 'F'
            $row->appendCell('E')->withValue('bar');

            // append multiple next cells after F (G,H,I), and assign values
            $row->appendCell('F', 3)->withValues([1, 2, 3]);

            return $row;
        });

        $excel->save();

        $cells = $this->read();

        $this->assertSame('foo', $cells['A']['v']);
        $this->assertSame(123, $cells['B']['v']);

        $this->assertSame('X', $cells['C']['v']);

        $this->assertSame('X', $cells['D']['v']);
        $this->assertSame('X', $cells['E']['v']);

        $this->assertSame('bar', $cells['F']['v']);

        $this->assertSame(1, $cells['G']['v']);
        $this->assertSame(2, $cells['H']['v']);
        $this->assertSame(3, $cells['I']['v']);
    }

    /**
     * The iterator must walk the whole row even if some cell holds a falsy value:
     * checking the current *value* instead of the key used to stop the loop in the
     * middle of the row, so every cell after it was silently lost on insertRow().
     */
    public function testIterationIsNotStoppedByFalsyCell()
    {
        $row = new RowTemplate();
        $row->addCell('A', 'first');
        $row->addCell('B', null);   // not scalar -> stored as is
        $row->addCell('C', []);     // empty array -> falsy too
        $row->addCell('D', 'last');

        $visited = [];
        foreach ($row as $colLetter => $cell) {
            $visited[] = $colLetter;
        }

        $this->assertSame(['A', 'B', 'C', 'D'], $visited);
    }

    /**
     * The same row, written to a real file: cells after the falsy one must not disappear.
     */
    public function testFalsyCellDoesNotTruncateInsertedRow()
    {
        $excel = $this->template();
        $sheet = $excel->sheet();

        $row = new RowTemplate();
        $row->addCell('A', 'first');
        $row->addCell('B', null);
        $row->addCell('C', 'last');

        $sheet->insertRow($row);
        $excel->save();

        $cells = $this->read();

        $this->assertSame('first', $cells['A']['v']);
        $this->assertArrayHasKey('C', $cells);
        $this->assertSame('last', $cells['C']['v']);
    }

    public function testSetValueRejectsNonScalar()
    {
        $this->expectException(Exception::class);
        $row = new RowTemplate();
        $row->setValue('A', ['not', 'scalar']);
    }

    public function testWithValueWithoutLastAddedThrows()
    {
        $this->expectException(Exception::class);
        $row = new RowTemplate();
        $row->withValue('x');
    }

    public function testWithValuesWithoutLastAddedThrows()
    {
        $this->expectException(Exception::class);
        $row = new RowTemplate();
        $row->withValues(['x']);
    }


}