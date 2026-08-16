<?php

declare(strict_types=1);

namespace avadim\FastExcelTemplator;

use avadim\FastExcelHelper\Helper;
use PHPUnit\Framework\TestCase;

/**
 * The lexer decides what inside a formula is a cell reference. Everything it recognizes is
 * rewritten to RC notation, everything it misses stays in A1 and silently stops following its row.
 * These tests therefore assert the recognition itself, not the arithmetic around it.
 */
final class FormulaLexerTest extends TestCase
{
    /**
     * All references the lexer recognizes in a formula, in order
     */
    private function refs(string $formula): array
    {
        $found = [];
        FormulaLexer::replaceReferences($formula, static function (string $ref) use (&$found) {
            $found[] = $ref;

            return $ref;
        });

        return $found;
    }

    public function recognitionProvider(): array
    {
        return [
            'plain references'        => ['=A1+B1', ['A1', 'B1']],
            'range is one token'      => ['=SUM(A1:A10)', ['A1:A10']],
            'absolute reference'      => ['=$A$1*C7', ['$A$1', 'C7']],
            'mixed reference'         => ['=A$1+$B2', ['A$1', '$B2']],
            'error constant'          => ['=SUM(#REF!)+A3', ['A3']],
            'error inside a string'   => ['=IF(A1>0,"#REF! and A1","no")&B1', ['A1', 'B1']],
            'doubled quotes'          => ['=A1&"he said ""A2"""', ['A1']],
            'exponent is a number'    => ['=1E3+E3', ['E3']],
            'sheet prefix'            => ['=SUM(Sheet2!B2:B99)', ['B2:B99']],
            'quoted sheet name'       => ["=SUM('Sheet two'!B2:B9)", ['B2:B9']],
            'non-ascii sheet name'    => ['=Лист2!A1+B2', ['A1', 'B2']],
            'external workbook'       => ['=[1]Sheet1!A1', ['A1']],
            'defined name'            => ['=TestDate1+1', []],
            'function name with digit' => ['=LOG10(A1)', ['A1']],
            'table reference'         => ['=SUMIF(Table1[Qty],">0")', []],
            'spilled range kept'      => ['=SUM(A1#)', []],
            'whole column kept'       => ['=SUM(A:A)', []],
            'last cell of a sheet'    => ['=XFD1048576', ['XFD1048576']],
            'outside a sheet is a name' => ['=ZZZ9999999+1', []],
            'nested calls'            => ['=IFERROR(VLOOKUP(A2,Sheet2!$A$1:$D$500,3,FALSE),"")', ['A2', '$A$1:$D$500']],
        ];
    }

    /**
     * @dataProvider recognitionProvider
     */
    public function testRecognizesReferences(string $formula, array $expected)
    {
        self::assertSame($expected, $this->refs($formula));
    }

    /**
     * Converting with one base and reading back with another must shift every relative
     * reference by exactly that difference — this is what re-basing an inserted row relies on.
     */
    public function testRelativeReferencesFollowTheRow()
    {
        $rc = FormulaLexer::convertA1toRC('=A1+$A$1+SUM(B2:B10)', 'C7');

        // read back one row lower
        $shifted = preg_replace_callback('/R\[?-?\d*]?C\[?-?\d*]?(?::R\[?-?\d*]?C\[?-?\d*]?)?/', static function ($m) {
            return Helper::RCtoA1($m[0], 'C8');
        }, $rc);

        // relative references moved one row down, the absolute one did not
        self::assertSame('=A2+$A$1+SUM(B3:B11)', $shifted);
    }

    public function testConversionAlwaysStartsWithEqualSign()
    {
        self::assertSame('=R[-6]C[-2]', FormulaLexer::convertA1toRC('A1', 'C7'));
        self::assertSame('=R[-6]C[-2]', FormulaLexer::convertA1toRC('=A1', 'C7'));
    }

    public function testFormulaWithoutReferencesIsUntouched()
    {
        self::assertSame('=TODAY()', FormulaLexer::convertA1toRC('=TODAY()', 'C7'));
        self::assertSame('=1+2*3', FormulaLexer::convertA1toRC('=1+2*3', 'C7'));
    }
}
