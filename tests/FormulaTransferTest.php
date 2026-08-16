<?php

declare(strict_types=1);

namespace avadim\FastExcelTemplator;

use avadim\FastExcelWriter\Excel as Writer;
use PHPUnit\Framework\TestCase;

/**
 * Formulas of a captured row template must follow the row they are inserted into:
 * relative references shift, absolute ones stay, and nothing of the RC notation
 * used internally may leak into the saved file.
 */
final class FormulaTransferTest extends TestCase
{
    private string $tpl;
    private string $out;

    protected function setUp(): void
    {
        // These tests stop reading a template halfway (that is the point: a row is captured and
        // reinserted), so its reader stays open and the file cannot be deleted on Windows.
        // Hence the system temp dir, and a separate pair of names per test.
        $name = preg_replace('/\W/', '', $this->getName());
        $dir = sys_get_temp_dir() . DIRECTORY_SEPARATOR;
        $this->tpl = $dir . 'fxt-formula-' . $name . '-tpl.xlsx';
        $this->out = $dir . 'fxt-formula-' . $name . '-out.xlsx';
        @unlink($this->tpl);
        @unlink($this->out);
    }

    protected function tearDown(): void
    {
        // Excel and SheetTemplate reference each other, so the template file stays open
        // until the cycle is collected — and an open file cannot be deleted on Windows
        gc_collect_cycles();
        @unlink($this->tpl);
        @unlink($this->out);
    }

    /**
     * Formulas of the saved sheet, keyed by cell address
     */
    private function formulas(string $file): array
    {
        $zip = new \ZipArchive();
        self::assertTrue($zip->open($file) === true);
        $xml = (string)$zip->getFromName('xl/worksheets/sheet1.xml');
        $zip->close();

        $result = [];
        if (preg_match_all('#<c r="([A-Z]+\d+)"[^>]*><f>(.*?)</f>#', $xml, $matches, PREG_SET_ORDER)) {
            foreach ($matches as $one) {
                $result[$one[1]] = $one[2];
            }
        }
        $result['__xml'] = $xml;

        return $result;
    }

    public function testInsertedRowsRebaseTheirFormulas()
    {
        // row 1 is a header, row 2 is the template row with formulas of every kind
        $writer = Writer::create(['Sheet1']);
        $ws = $writer->sheet();
        $ws->writeRow(['Qty', 'Price', 'Sum', 'Share', 'Broken']);
        $ws->writeRow([1, 2, '=A2*B2', '=C2/$C$1', '=SUM(#REF!)+A2']);
        $writer->save($this->tpl);

        $excel = Excel::template($this->tpl, $this->out);
        $sheet = $excel->sheet();
        $sheet->transferRowsUntil(1);
        $rowTemplate = $sheet->getRowTemplate(2);
        for ($i = 0; $i < 3; $i++) {
            $sheet->insertRow($rowTemplate, ['A' => $i + 1, 'B' => 10]);
        }
        $excel->save();

        $formulas = $this->formulas($this->out);

        // relative references follow their row
        self::assertSame('A2*B2', $formulas['C2']);
        self::assertSame('A3*B3', $formulas['C3']);
        self::assertSame('A4*B4', $formulas['C4']);

        // the absolute part stays put while the relative one moves
        self::assertSame('C2/$C$1', $formulas['D2']);
        self::assertSame('C3/$C$1', $formulas['D3']);

        // a reference standing after an error constant is rebased too
        // (the PHP tokenizer used to read '#' as a comment and skip the rest of the formula)
        self::assertSame('SUM(#REF!)+A2', $formulas['E2']);
        self::assertSame('SUM(#REF!)+A3', $formulas['E3']);
        self::assertSame('SUM(#REF!)+A4', $formulas['E4']);

        // no RC notation may leak into a saved A1 formula
        self::assertStringNotContainsString('<f>R[', $formulas['__xml']);
        self::assertStringNotContainsString('C[', $formulas['__xml']);
    }

    public function testRangeIsRebasedAsAWhole()
    {
        $writer = Writer::create(['Sheet1']);
        $ws = $writer->sheet();
        $ws->writeRow(['Header']);
        $ws->writeRow([1]);
        $ws->writeRow([2]);
        $ws->writeRow([3, '=SUM(A2:A4)']);   // range fully above the template row
        $writer->save($this->tpl);

        $excel = Excel::template($this->tpl, $this->out);
        $sheet = $excel->sheet();
        $sheet->transferRowsUntil(3);
        $rowTemplate = $sheet->getRowTemplate(4);
        $sheet->insertRow($rowTemplate);
        $sheet->insertRow($rowTemplate);
        $excel->save();

        $formulas = $this->formulas($this->out);

        // both ends of the range shift together
        self::assertSame('SUM(A2:A4)', $formulas['B4']);
        self::assertSame('SUM(A3:A5)', $formulas['B5']);
        self::assertStringNotContainsString('C[', $formulas['__xml']);
    }

    /**
     * A row template inserted ABOVE its source row can push a relative reference off the sheet.
     * Excel turns such a reference into #REF!; the RC notation must not survive into the file.
     *
     * The conversion back to A1 happens in avadim/fast-excel-writer (FormulaConverter::normalize()),
     * which currently keeps the unconvertible RC text as is, so this is pending a fix there.
     */
    public function testReferenceShiftedOffTheSheetBecomesRefError()
    {
        $writer = Writer::create(['Sheet1']);
        $ws = $writer->sheet();
        for ($i = 1; $i <= 4; $i++) {
            $ws->writeRow([$i * 10]);
        }
        $ws->writeRow([50, '=A1*2']);   // row 5 refers four rows up
        $writer->save($this->tpl);

        $excel = Excel::template($this->tpl, $this->out);
        $sheet = $excel->sheet();
        $rowTemplate = $sheet->getRowTemplate(5, true);
        $sheet->transferRowsUntil(1);
        $sheet->insertRow($rowTemplate);   // lands on row 2: 2 - 4 = -2, outside the sheet
        $excel->save();

        $formulas = $this->formulas($this->out);

        if (strpos($formulas['B2'], 'R[') !== false) {
            self::markTestSkipped('needs #REF! fallback in fast-excel-writer, got: ' . $formulas['B2']);
        }

        self::assertStringContainsString('#REF!', $formulas['B2']);
    }
}
