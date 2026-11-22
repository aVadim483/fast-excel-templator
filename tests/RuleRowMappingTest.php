<?php

declare(strict_types=1);

namespace avadim\FastExcelTemplator;

use avadim\FastExcelReader\Excel as Reader;
use avadim\FastExcelWriter\DataValidation\DataValidation;
use avadim\FastExcelWriter\Excel as Writer;
use PHPUnit\Framework\TestCase;

final class RuleRowMappingTest extends TestCase
{
    private string $template;
    private string $output;

    protected function setUp(): void
    {
        $this->template = sys_get_temp_dir() . '/fxt-rule-' . uniqid('', true) . '.xlsx';
        $this->output = $this->template . '-out.xlsx';
    }

    protected function tearDown(): void
    {
        gc_collect_cycles();
        @unlink($this->template);
        @unlink($this->output);
    }

    /** @param array<string, array<string, string>> $rules Formulas by area, grouped by sheet. */
    private function makeTemplate(array $rules, array $missingRows = []): void
    {
        $writer = Writer::create(array_keys($rules));
        foreach (array_keys($rules) as $sheet) {
            for ($i = 1; $i <= 7; $i++) {
                $writer->sheet($sheet)->writeRow([$i, '=A' . $i]);
            }
        }
        self::assertTrue($writer->save($this->template));
        $zip = new \ZipArchive();
        self::assertTrue($zip->open($this->template));
        foreach (array_values($rules) as $index => $sheetRules) {
            $entry = 'xl/worksheets/sheet' . ($index + 1) . '.xml';
            $xml = $zip->getFromName($entry);
            foreach ($missingRows as $row) {
                $xml = preg_replace('/<row\b[^>]*\br="' . $row . '"[^>]*>.*?<\/row>/s', '', $xml);
            }
            $conditions = '';
            $validations = '<dataValidations count="' . count($sheetRules) . '" xWindow="23">';
            $priority = 0;
            foreach ($sheetRules as $area => $formula) {
                $formula = htmlspecialchars($formula, ENT_XML1);
                $conditions .= '<conditionalFormatting sqref="' . $area . '"><cfRule type="expression" priority="'
                    . ++$priority . '"><formula>' . $formula . '</formula></cfRule></conditionalFormatting>';
                $validations .= '<dataValidation sqref="' . $area . '" type="custom" errorTitle="Keep me">'
                    . '<formula1>' . $formula . '</formula1></dataValidation>';
            }
            self::assertTrue($zip->addFromString($entry, str_replace(
                '</sheetData>', '</sheetData>' . $conditions . $validations . '</dataValidations>', $xml
            )));
        }
        self::assertTrue($zip->close());
    }

    private function xpath(int $sheet): \DOMXPath
    {
        $zip = new \ZipArchive();
        self::assertTrue($zip->open($this->output));
        $xml = $zip->getFromName('xl/worksheets/sheet' . $sheet . '.xml');
        $zip->close();
        $dom = new \DOMDocument();
        self::assertTrue($dom->loadXML($xml));
        $xpath = new \DOMXPath($dom);
        $xpath->registerNamespace('s', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
        return $xpath;
    }

    private function rules(\DOMXPath $xpath, string $query): array
    {
        $rules = [];
        foreach ($xpath->query($query) as $node) {
            $rules[$node->getAttribute('sqref')] = trim($node->textContent);
        }
        return $rules;
    }

    public function testStructuralOperationsRemapRulesIndependentlyOnTwoSheets(): void
    {
        $this->makeTemplate([
            'Alpha' => ['B3:B5' => '1', 'C2' => '2', 'D1:D3 F6 G7:H7' => '$A$1>0'],
            'Beta' => ['P2' => '3', 'Q3:Q4' => '4', 'R5:R7' => '5'],
        ]);
        $excel = Excel::template($this->template, $this->output, ['strict_metadata' => true]);
        $alpha = $excel->sheet('Alpha');
        $alpha->transferRowsUntil(1)->skipRows(1)->transferRowsUntil(3)->insertRow(['A' => 90]);
        $alpha->transferRowsUntil(5)->insertRow(['A' => 91])->transferRows();
        $alpha->sheetWriter->addDataValidation('A9', DataValidation::dropDown(['Yes', 'No']));

        $beta = $excel->sheet('Beta');
        $beta->transferRowsUntil(1);
        $sample = $beta->getRowTemplate(2);
        $beta->insertRow($sample)->insertRow($sample);
        $beta->rows(static function ($source, $target, $row) {
            if ($source === 3) {
                return null;
            }
            return $source === 5 ? false : $row;
        });

        self::assertTrue($excel->save());
        self::assertSame([], $excel->metadataDiagnostics());
        $alphaXml = $this->xpath(1);
        $alphaRules = ['B2 B4:B5' => '1', 'D1:D2 F7 G8:H8' => '$A$1>0'];
        self::assertSame($alphaRules, $this->rules($alphaXml, '//s:conditionalFormatting'));
        self::assertSame($alphaRules + ['A9' => '"Yes,No"'], $this->rules($alphaXml, '//s:dataValidation'));
        self::assertSame('3', $alphaXml->evaluate('string(//s:dataValidations/@count)'));
        self::assertSame('23', $alphaXml->evaluate('string(//s:dataValidations/@xWindow)'));
        self::assertSame('Keep me', $alphaXml->evaluate('string(//s:dataValidation/@errorTitle)'));
        self::assertSame('A2', $alphaXml->evaluate('string(//s:c[@r="B2"]/s:f)'));
        $betaXml = $this->xpath(2);
        $betaRules = ['P2:P3' => '3', 'Q4' => '4'];
        self::assertSame($betaRules, $this->rules($betaXml, '//s:conditionalFormatting'));
        self::assertSame($betaRules, $this->rules($betaXml, '//s:dataValidation'));
        self::assertSame('2', $betaXml->evaluate('string(//s:dataValidations/@count)'));
        $output = Reader::open($this->output);
        self::assertSame([1, 3, 90, 4, 5, 91, 6, 7], array_column($output->sheet('Alpha')->readRows(), 'A'));
        self::assertSame([1, 2, 2, 4], array_column($output->sheet('Beta')->readRows(), 'A'));
    }

    /** @dataProvider relativeRuleFormulas */
    public function testRelativeRuleFormulasFollowUniformShiftAcrossTransferCalls(string $formula, string $shifted): void
    {
        $this->makeTemplate(['Data' => ['B3:B5' => $formula]]);
        $excel = Excel::template($this->template, $this->output, ['strict_metadata' => true]);
        $sheet = $excel->sheet();
        $sheet->transferRowsUntil(2, static function ($source, $target, $row) {
            return $source === 2 ? false : $row;
        });
        $sheet->transferRows();
        self::assertTrue($excel->save());
        $xpath = $this->xpath(1);
        $expected = ['B2:B4' => $shifted];
        self::assertSame($expected, $this->rules($xpath, '//s:dataValidation'));
        self::assertSame($expected, $this->rules($xpath, '//s:conditionalFormatting'));
        self::assertSame([], $excel->metadataDiagnostics());
    }

    public static function relativeRuleFormulas(): array
    {
        return [
            'comparison' => ['$A3>A$1', '$A2>A$1'],
            'logical functions' => [
                'AND($A3<>"OR(A3)",oR($B3>0,NOT($C$1)))',
                'AND($A2<>"OR(A3)",oR($B2>0,NOT($C$1)))',
            ],
        ];
    }

    public function testContextDependentRuleFunctionIsRejectedAfterUniformShift(): void
    {
        $this->makeTemplate(['Data' => ['B3:B5' => 'AND(ROW()>2,A3>0)']]);
        $excel = Excel::template($this->template, $this->output, ['strict_metadata' => true]);
        $excel->sheet()->transferRowsUntil(1)->skipRows(1)->transferRows();
        $this->expectException(Exception::class);
        $this->expectExceptionMessage('on "Data" (sqref="B3:B5")');
        $excel->save();
    }

    public function testSplitWithRelativeFormulaIsDiagnosed(): void
    {
        $formula = 'AND(A1>0,A1<10)';
        $this->makeTemplate(['Data' => ['B1:B3' => $formula]]);
        foreach ([false, true] as $strict) {
            $excel = Excel::template($this->template, $this->output, ['strict_metadata' => $strict]);
            $excel->sheet()->transferRowsUntil(1)->insertRow(['A' => 90])->transferRows();
            if ($strict) {
                $this->expectException(Exception::class);
                $this->expectExceptionMessage('on "Data" (sqref="B1:B3")');
            }
            self::assertTrue($excel->save());
            self::assertSame('unmapped-rule-formula', $excel->metadataDiagnostics()[0]['code']);
            self::assertSame(['B1 B3:B4' => $formula], $this->rules($this->xpath(1), '//s:dataValidation'));
        }
    }

    public function testDirectWriterRowChangesAreRejectedInStrictMode(): void
    {
        $this->makeTemplate(['Data' => ['B1:B3' => '1']]);
        $excel = Excel::template($this->template, $this->output, ['strict_metadata' => true]);
        $excel->sheet()->sheetWriter->skipRow();
        $excel->sheet()->transferRows();
        $this->expectException(Exception::class);
        $this->expectExceptionMessage('Direct Writer row or cell writes');
        $excel->save();
    }

    public function testSparsePositionsAreMappedWithoutConsumingTheNextRow(): void
    {
        $this->makeTemplate(['Data' => ['B2:B4' => '1']], [2, 3]);
        $excel = Excel::template($this->template, $this->output, ['strict_metadata' => true]);
        $excel->sheet()->transferRowsUntil(2)->insertRow(['A' => 90])->skipRowsUntil(3)->transferRows();
        self::assertTrue($excel->save());
        $xpath = $this->xpath(1);
        self::assertSame(['B2 B4:B5' => '1'], $this->rules($xpath, '//s:dataValidation'));
        self::assertSame('4', $xpath->evaluate('string(//s:c[@r="A5"]/s:v)'));
    }

    public function testEmptyValidationContainerIsRemovedAndReplacementKeepsSourceRule(): void
    {
        $this->makeTemplate(['Alpha' => ['B2' => '1'], 'Beta' => ['B2' => '1']]);
        $excel = Excel::template($this->template, $this->output, ['strict_metadata' => true]);
        $excel->sheet('Alpha')->transferRowsUntil(1)->skipRows(1)->transferRows();
        $excel->sheet('Beta')->transferRowsUntil(1)->replaceRow(['A' => 90])->transferRows();
        self::assertTrue($excel->save());
        self::assertSame(0, $this->xpath(1)->query('//s:dataValidations|//s:conditionalFormatting')->length);
        self::assertSame(['B2' => '1'], $this->rules($this->xpath(2), '//s:dataValidation'));
    }
}
