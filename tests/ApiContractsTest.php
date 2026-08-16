<?php

declare(strict_types=1);

namespace avadim\FastExcelTemplator;

use PHPUnit\Framework\TestCase;

/**
 * Contracts of the public API that used to be broken or merely implied:
 * an Iterator that can actually be iterated, a collection usable without a sheet,
 * a subclassable entry point, a library that does not print, and one exception type.
 */
final class ApiContractsTest extends TestCase
{
    const TPL_FILE = __DIR__ . '/test_files/test.xlsx';

    private function collection(int $count): RowTemplateCollection
    {
        $rows = [];
        for ($i = 1; $i <= $count; $i++) {
            $rows[$i] = (new RowTemplate())->addCell('A', 'row ' . $i);
        }

        return new RowTemplateCollection($rows);
    }

    /**
     * Keys collected by a foreach, with a hard stop: before the fix the loop never ended,
     * and a test that hangs the suite is worse than one that fails
     */
    private function keysOf(RowTemplateCollection $collection, int $limit = 10): array
    {
        $keys = [];
        foreach ($collection as $key => $rowTemplate) {
            $keys[] = $key;
            if (count($keys) >= $limit) {
                self::fail('foreach over RowTemplateCollection does not stop');
            }
        }

        return $keys;
    }

    /**
     * foreach over the collection used to never end, because next() wraps around by design
     */
    public function testCollectionCanBeIterated()
    {
        self::assertSame([1, 2, 3], $this->keysOf($this->collection(3)));
    }

    public function testCollectionCanBeIteratedTwice()
    {
        $collection = $this->collection(2);

        self::assertSame([1, 2], $this->keysOf($collection));
        self::assertSame([1, 2], $this->keysOf($collection), 'the second pass must repeat the first');
    }

    /**
     * next() must keep cycling: insertRow() takes the next template on every call
     * and starts over at the end (that is how alternating row styles work)
     */
    public function testNextKeepsCyclingThroughTemplates()
    {
        $collection = $this->collection(2);

        $values = [];
        for ($i = 0; $i < 5; $i++) {
            $values[] = $collection->next()->getValue('A');
        }

        self::assertSame(['row 1', 'row 2', 'row 1', 'row 2', 'row 1'], $values);
    }

    /**
     * A collection built by hand has no sheet; cloning a cell must still work
     * (copying the column width is the only part that needs the sheet)
     */
    public function testCloneCellWorksWithoutSheet()
    {
        $collection = $this->collection(1);
        $collection->cloneCell('A', 'B');

        $cells = $collection->current()->cells();

        self::assertArrayHasKey('B', $cells);
        self::assertSame('row 1', $collection->current()->getValue('B'));
    }

    /**
     * template() must return an instance of the called class, not of Excel
     */
    public function testTemplateIsSubclassable()
    {
        $excel = TemplatorSubclass::template(self::TPL_FILE);

        self::assertInstanceOf(TemplatorSubclass::class, $excel);
    }

    /**
     * A library method must return data, not echo it
     */
    public function testValidateReturnsFileListWithoutPrinting()
    {
        $reader = Excel::createReader(self::TPL_FILE);

        ob_start();
        $files = $reader->validate();
        $printed = ob_get_clean();

        self::assertSame('', $printed, 'validate() must not print anything');
        self::assertIsArray($files);
        self::assertNotEmpty($files);
    }

    /**
     * Row range errors are reported with the package exception; since it extends
     * RuntimeException, code catching the old type keeps working
     */
    public function testRowRangeErrorUsesPackageException()
    {
        $sheet = Excel::template(self::TPL_FILE)->sheet();

        try {
            $sheet->getRowTemplate(100500);
            self::fail('an out-of-range row number must be rejected');
        }
        catch (Exception $e) {
            self::assertInstanceOf(\RuntimeException::class, $e, 'must stay catchable as before');
        }
    }
}

/**
 * Used by testTemplateIsSubclassable()
 */
class TemplatorSubclass extends Excel
{
}
