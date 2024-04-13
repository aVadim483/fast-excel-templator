<?php

use avadim\FastExcelTemplator\Excel;

require_once __DIR__ . '/../vendor/autoload.php';
require_once __DIR__ . '/../src/autoload.php';
$tpl = __DIR__ . '/files/demo1-tpl.xlsx';
$out = __DIR__ . '/files/demo1-out.xlsx';

$time = microtime(true);

// Open template and set output file
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

// Replacement strings and substrings on the sheet
$sheet
    ->fill($fillData)
    ->replace($replaceData)
;

// Transfer rows 1-6 from templates to output file
$sheet->transferRowsUntil(6);

// Get the specified row (number 7) as a template and go to the next row in the template
$rowTemplate = $sheet->getRowTemplate(7);
$count = 0;
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

$sheet->transferRows();

$excel->save();

echo 'Output file: ' . $out . '<br>';
echo 'Elapsed time: ' . round(microtime(true) - $time, 3) . ' sec';