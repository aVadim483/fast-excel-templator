<?php

declare(strict_types=1);

namespace avadim\FastExcelTemplator;

use avadim\FastExcelTemplator\Excel;
use PHPUnit\Framework\TestCase;

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

    public function test02()
    {
        $tpl = self::DEMO_DIR . 'demo1-tpl.xlsx';
        $out = __DIR__ . '/demo1-out.xlsx';

        $excel = Excel::template($tpl, $out);
        $sheet = $excel->sheet();

        $sheet->sheetWriter
            ->setColWidth(['C', 'D'], 250);
        $sheet->transferRows();
        $excel->save();
        $this->assertTrue(is_file($out));

        $excelReader = \avadim\FastExcelReader\Excel::open($out);
        $sheet = $excelReader->sheet();
        $colAttributes = $sheet->getColAttributes();
        $this->assertEquals(250, (int)$colAttributes['C']['width']);
        $this->assertEquals(250, (int)$colAttributes['D']['width']);

        unlink($out);
    }

    public function test03()
    {
        $tpl = __DIR__ . '/files/test-formulas.xlsx';
        $out = __DIR__ . '/files/test-formulas-out.xlsx';

        $excel = Excel::template($tpl, $out);

        foreach ($excel->sheets() as $sheet) {
            $sheet->transferRows();
        }

        $excel->save();
        $this->assertTrue(is_file($out));

        $excelReader = \avadim\FastExcelReader\Excel::open($out);
        $excelReader->dateFormatter(true);
        $cells = $excelReader->readCellsWithStyles();

        $c1 = $cells['B2'];unset($c1['s']);
        $c0 = ['v' => '23.01.1985', 'f' => null, 'o' => '31070', 't' => 'date'];
        $this->assertEquals($c0, $c1);

        $c1 = $cells['C2'];unset($c1['s']);
        $c0 = ['v' => '=B2+1', 'f' => '=B2+1', 'o' => '=B2+1', 't' => 'date'];
        $this->assertEquals($c0, $c1);

        $c1 = $cells['C3'];unset($c1['s']);
        $c0 = ['v' => '=Sheet2!B2', 'f' => '=Sheet2!B2', 'o' => '=Sheet2!B2', 't' => 'date'];
        $this->assertEquals($c0, $c1);

        $c1 = $cells['C6'];unset($c1['s']);
        $c0 = ['v' => '=TestDate1+1', 'f' => '=TestDate1+1', 'o' => '=TestDate1+1', 't' => 'date'];
        $this->assertEquals($c0, $c1);

        $c1 = $cells['C9'];unset($c1['s']);
        $c0 = ['v' => '="qwe" & TestDate1', 'f' => '="qwe" & TestDate1', 'o' => '="qwe" & TestDate1', 't' => ''];
        $this->assertEquals($c0, $c1);

        $c1 = $cells['C10'];unset($c1['s']);
        $c0 = ['v' => '=SUM(C2:C9)', 'f' => '=SUM(C2:C9)', 'o' => '=SUM(C2:C9)', 't' => 'number'];
        $this->assertEquals($c0, $c1);

        unlink($out);
    }
}

