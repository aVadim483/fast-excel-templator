<?php

declare(strict_types=1);

namespace avadim\FastExcelTemplator;

use PHPUnit\Framework\TestCase;

final class StyleTest extends TestCase
{
    const INP_FILE = __DIR__ . '/test_files/test.xlsx';
    const OUT_FILE = __DIR__ . '/test_files/test-style-out.xlsx';

    protected function setUp(): void
    {
        if (is_file(self::OUT_FILE)) {
            unlink(self::OUT_FILE);
        }
    }

    protected function template(): Excel
    {
        return Excel::template(self::INP_FILE, self::OUT_FILE);
    }

    public function testSetRowStyle()
    {
        $excel = $this->template();
        $sheet = $excel->sheet();

        $style = [
            'fill' => [
                'fill-color' => '#ff0000',
            ],
            'font' => [
                'font-style' => 'bold',
            ],
        ];

        $sheet->rows(function ($sourceRowNum, $targetRowNum, RowTemplate $row) use ($style) {
            if ($sourceRowNum === 1) {
                $row->setRowStyle($style);
                $row->setValue('A', 'Style Row 1');
            }
            return $row;
        });

        $excel->save();

        $this->assertTrue(is_file(self::OUT_FILE));

        $excelReader = \avadim\FastExcelReader\Excel::open(self::OUT_FILE);
        $sheetReader = $excelReader->sheet();

        // В fast-excel-reader атрибуты строки можно получить, используя RESULT_MODE_ROW
        $rows = [];
        foreach ($sheetReader->nextRow([], \avadim\FastExcelReader\Excel::RESULT_MODE_ROW) as $rNum => $rData) {
            $rows[$rNum] = $rData;
        }

        $this->assertArrayHasKey(1, $rows);
        $styleIdx = (int)$rows[1]['__row']['s'];
        
        $styles = $excelReader->readStyles();
        $this->assertArrayHasKey($styleIdx, $styles['cellXfs']);
        
        $xf = $styles['cellXfs'][$styleIdx];
        $fill = $styles['fills'][$xf['fillId']];
        
        $this->assertEquals('#FF0000', $fill['fill-color']);
        $font = $styles['fonts'][$xf['fontId']];
        $this->assertNotEmpty($font['font-style-bold']);
    }

    public function testSetColStyle()
    {
        $excel = $this->template();
        $sheet = $excel->sheet();
        
        $writer = $sheet->sheetWriter;

        $style = [
            'fill' => [
                'fill-color' => '#00ff00',
            ],
        ];

        $writer->setColStyle('B', $style);

        $sheet->rows(function ($sourceRowNum, $targetRowNum, RowTemplate $row) {
            return $row;
        });

        $excel->save();

        $this->assertTrue(is_file(self::OUT_FILE));

        $excelReader = \avadim\FastExcelReader\Excel::open(self::OUT_FILE);
        $sheetReader = $excelReader->sheet();
        $colAttributes = $sheetReader->getColAttributes();
        
        $this->assertArrayHasKey('B', $colAttributes);
        $styleIdx = (int)$colAttributes['B']['style'];
        $styles = $excelReader->readStyles();
        $this->assertArrayHasKey($styleIdx, $styles['cellXfs']);
        
        $xf = $styles['cellXfs'][$styleIdx];
        $fill = $styles['fills'][$xf['fillId']];

        $this->assertEquals('#00FF00', $fill['fill-color']);
    }
}
