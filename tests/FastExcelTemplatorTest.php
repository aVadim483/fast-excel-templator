<?php

declare(strict_types=1);

namespace avadim\FastExcelTemplator;

use PHPUnit\Framework\TestCase;
use avadim\FastExcelTemplator\Excel;

final class FastExcelTemplatorTest extends TestCase
{
    const DEMO_DIR = __DIR__ . '/../demo/files/';

    public function test01()
    {
        // =====================
        $tpl = self::DEMO_DIR . 'demo1-tpl.xlsx';
        $out = __DIR__ . '/demo1-out.xlsx';

        $excel = Excel::template($tpl, $out);
        $sheet = $excel->sheet();

        $fillData = [
            '{{COMPANY}}' => 'Comp Stock Shop',
            '{{ADDRESS}}' => '123 ABC Street',
            '{{CITY}}' => 'Peace City, TN',
        ];
        $replaceData = ['{{BULK_QTY}}' => 12, '{{DATE}}' => date('m/d/Y')];
        $list = [
            [
                'number' => 'AA-8465-947563',
                'description' => 'NEC Mate Type ML PC-MK29MLZD1FWG (Corei5-3470S/2GB/250GB/Multi/No OF/Win 11 Home)',
                'price1' => 816,
                'price2' => 683,
            ],
            [
                'number' => 'QR-7956-048914',
                'description' => 'Acer Aspire TC-1760-UA92 (Core i5-12400/12GB 3200MHz DDR4/512GB SSD/Win 11 Home)',
                'price1' => 414,
                'price2' => 339,
            ],
            [
                'number' => 'BD-3966-285936',
                'description' => 'Dell Optiplex 7050 SFF (Core i7-7700/32GB DDR4/1TB SSD/Win 10 Pro/Renewed)',
                'price1' => 249,
                'price2' => 198,
            ],
        ];

        $sheet
            ->fill($fillData)
            ->replace($replaceData)
        ;
        $sheet->transferRowsUntil(6);
        $rowTemplate = $sheet->getRowTemplate(7);
        foreach ($list as $record) {
            $rowData = [
                // In the column A wil be written value from field 'number'
                'A' => $record['number'],
                // In the column B wil be written value from field 'description'
                'B' => $record['description'],
                // And so on...
                'C' => $record['price1'],
                'D' => $record['price2'],
            ];
            $sheet->insertRow($rowTemplate, $rowData);
        }
        $excel->save();
        $this->assertTrue(is_file($out));

        $excelReader = \avadim\FastExcelReader\Excel::open($out);
        $cells = $excelReader->readCells();

        $this->assertEquals($fillData['{{COMPANY}}'], $cells['A2']);
        $this->assertEquals('*Bulk pricing applies to quantities of 12 or more', $cells['D4']);
        $this->assertEquals(249, $cells['C9']);

        unlink($out);
    }

}

