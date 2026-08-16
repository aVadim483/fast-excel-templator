<?php

declare(strict_types=1);

namespace avadim\FastExcelTemplator;

use avadim\FastExcelWriter\Excel as Writer;
use PHPUnit\Framework\TestCase;

/**
 * Regression tests for defects found while auditing the library:
 *  - the relationship to the dropped calcChain.xml was left dangling (wrong part name in ExcelWriter)
 *  - outputFile() crashed when the output file was not passed to template()
 *  - a cell without the 't' key raised a warning, and a date cell was written twice
 */
final class RegressionFixesTest extends TestCase
{
    const TPL_FORMULAS = __DIR__ . '/test_files/test-formulas.xlsx';

    private function zipEntry(string $file, string $entry): string
    {
        $zip = new \ZipArchive();
        self::assertTrue($zip->open($file) === true);
        $content = $zip->getFromName($entry);
        $zip->close();

        return (string)$content;
    }

    /**
     * When calcChain.xml is removed from the output, the relationship pointing to it
     * must be removed too, otherwise the workbook keeps a reference to a missing part.
     */
    public function testCalcChainRelationshipIsRemoved()
    {
        $out = __DIR__ . '/test_files/test-calcchain-out.xlsx';
        @unlink($out);

        // the template really contains calcChain.xml and a relationship to it
        self::assertStringContainsString('calcChain', $this->zipEntry(self::TPL_FORMULAS, 'xl/_rels/workbook.xml.rels'));

        $excel = Excel::template(self::TPL_FORMULAS, $out);
        $excel->sheet()->transferRows();
        $excel->save();

        $zip = new \ZipArchive();
        self::assertTrue($zip->open($out) === true);
        $calcChainExists = (bool)$zip->statName('xl/calcChain.xml');
        $zip->close();

        // the part itself is dropped...
        self::assertFalse($calcChainExists);
        // ...so neither the content types nor the relationships may still reference it
        self::assertStringNotContainsString('calcChain', $this->zipEntry($out, '[Content_Types].xml'));
        self::assertStringNotContainsString('calcChain', $this->zipEntry($out, 'xl/_rels/workbook.xml.rels'));

        @unlink($out);
    }

    /**
     * template() may be called without an output file (it can be given later to save()),
     * so outputFile() must not throw on an uninitialized typed property.
     */
    public function testOutputFileWithoutOutputArgument()
    {
        $excel = Excel::template(self::TPL_FORMULAS);

        self::assertSame('', $excel->outputFile());
    }

    /**
     * A cell array built by hand has no 't' key; writing such a row must not raise a warning.
     */
    public function testCellWithoutTypeKeyIsWritten()
    {
        $tpl = __DIR__ . '/test_files/test-notype-tpl.xlsx';
        $out = __DIR__ . '/test_files/test-notype-out.xlsx';
        @unlink($tpl);
        @unlink($out);

        $writer = Writer::create(['Sheet1']);
        $writer->sheet()->writeRow(['a', 'b']);
        $writer->save($tpl);

        $excel = Excel::template($tpl, $out);
        $excel->sheet()->rows(function ($sourceRowNum, $targetRowNum, RowTemplate $row) {
            // no 't', no 's' — only a value
            $row->addCell('C', ['v' => 'no-type']);

            return $row;
        });
        $excel->save();

        $cells = \avadim\FastExcelReader\Excel::open($out)->readRows()[1];
        self::assertSame('no-type', $cells['C']);

        @unlink($tpl);
        @unlink($out);
    }

    /**
     * A date cell must keep its value and its date format after being transferred.
     * (It used to be written twice: once with the format, then again without it.)
     */
    public function testDateCellKeepsValueAndFormat()
    {
        $tpl = __DIR__ . '/test_files/test-date-tpl.xlsx';
        $out = __DIR__ . '/test_files/test-date-out.xlsx';
        @unlink($tpl);
        @unlink($out);

        $writer = Writer::create(['Sheet1']);
        $writer->sheet()->writeRow(['2026-08-16', 'text'], ['format' => '@date']);
        $writer->save($tpl);

        $excel = Excel::template($tpl, $out);
        $excel->sheet()->transferRows();
        $excel->save();

        $sourceCell = \avadim\FastExcelReader\Excel::open($tpl)->readCellsWithStyles()['A1'];
        $targetCell = \avadim\FastExcelReader\Excel::open($out)->readCellsWithStyles()['A1'];

        // the underlying date serial survives the transfer (locale-independent, unlike 'v')
        self::assertSame($sourceCell['o'], $targetCell['o']);
        // and the cell is still recognized as a date, i.e. it kept its date number format
        self::assertSame('date', $targetCell['t']);
        self::assertSame(
            $sourceCell['s']['format']['format-pattern'],
            $targetCell['s']['format']['format-pattern']
        );

        @unlink($tpl);
        @unlink($out);
    }
}
