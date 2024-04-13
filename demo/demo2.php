<?php

use avadim\FastExcelTemplator\Excel;

require_once __DIR__ . '/../vendor/autoload.php';
require_once __DIR__ . '/../src/autoload.php';

$tpl = __DIR__ . '/files/demo2-tpl.xlsx';
$out = __DIR__ . '/files/demo2-out.xlsx';

$time = microtime(true);

// Open template and set output file
$excel = Excel::template($tpl);
$sheet = $excel->sheet();

$data = [
    'Aberdeen City' => [
        'January' => 137145,
        'February' => 120000,
        'March' => 140000,
        'April' => 134995,
        'May' => 147500,
        'June' => 150000,
    ],
    'North Ayrshire' => [
        'January' => 114500,
        'February' => 110000,
        'March' => 135000,
        'April' => 100000,
        'May' => 127500,
        'June' => 125000,
    ],
    'Moray' => [
        'January' => 180000,
        'February' => 182625,
        'March' => 195250,
        'April' => 180000,
        'May' => 188000,
        'June' => 194000,
    ],
];

// Transfer rows 1-3 from templates to output file
$sheet->transferRowsUntil(3);

$head = $sheet->getRowTemplate(4);
$head->cloneCell('D', 'E-I');
$sheet->insertRow($head, ['D' => 'January', 'E' => 'February', 'F' => 'March', 'G' => 'April', 'H' => 'May', 'I' => 'June']);

$rowTemplate = $sheet->getRowTemplate(5);
$rowTemplate->cloneCell('D', ['E', 'F', 'G-I']);
$cnt = 0;
foreach ($data as $locName => $locData) {
    $rowData = [
        'B' => ++$cnt,
        'C' => $locName,
        'D' => $locData['January'],
        'E' => $locData['February'],
        'F' => $locData['March'],
        'G' => $locData['April'],
        'H' => $locData['May'],
        'I' => $locData['June'],
    ];
    $sheet->insertRow($rowTemplate, $rowData);
}

$excel->save($out);

echo 'Output file: ' . $out . '<br>';
echo 'Elapsed time: ' . round(microtime(true) - $time, 3) . ' sec';
