<?php

use avadim\FastExcelTemplator\Excel;

require_once __DIR__ . '/../vendor/autoload.php';
require_once __DIR__ . '/../src/autoload.php';

$tpl = __DIR__ . '/files/demo3-tpl.xlsx';
$out = __DIR__ . '/files/demo3-out.xlsx';

$time = microtime(true);

// Open template and set output file
$excel = Excel::template($tpl);
$excel->dateFormatter(true);
$sheet = $excel->sheet();

$invoiceNo = time();
$invoiceDate = date('Y-m-d');
$data = [
    '{{INVOICE_NO}}' => $invoiceNo,
    '{{INVOICE_DATE}}' => $invoiceDate,

    '{{COMPANY_NAME}}' => 'Tech Prod Corp.',
    '{{STREET_ADDRESS}}' => '123 ABC Street',
    '{{CITY_ADDRESS}}' => 'Peace City',
    '{{PHONE_NUMBER}}' => '(555) 555-1234'		,
    '{{EMAIL_ADDRESS}}' => 'sales@tech-prod-corp.com',

    '{{DUE_IN_DAYS}}' => '5',

    '{{SHIPPING_NAME}}' => 'Joe Done',
    '{{SHIPPING_ADDRESS}}' => 'Chennai St, TN',
    '{{SHIPPING_PHONE}}' => '(444) 444-1234',

    '{{BILLING_ADDRESS}}' => 'Kanya Kumari, TN',
    '{{BILLING_PHONE}}' => '(333) 333-1234',
    '{{BILLING_EMAIL}}' => 'bill@tech-prod-corp.com',

    '{{REMARKS}}' => 'Some remarks are here',
];

$excel
    ->fill($data)
    ->replace(['{{INVOICE_NO}}' => $invoiceNo, '{{DUE_IN_DAYS}}' => '5'])
;

$sheet->transferRowsUntil(17);
$rows = $sheet->getRowTemplates(18, 21);

$firstRowNum = $sheet->lastWrittenRowNum() + 1;
for ($i = 1; $i <= 9; $i++) {
    $sheet->insertRow($rows, ['B' => 'item name ' . $i, 'D' => random_int(2, 20), 'E' => rand(1000, 9000) / 100]);
}
$lastRowNum = $sheet->lastWrittenRowNum();

$formula = '=SUM(F' . $firstRowNum . ':F' . $lastRowNum . ')';

$sheet->transferRows(null, function ($targetRowNum, $rowTemplate) use ($formula) {
    if ($rowTemplate->rowNumber() === 22) {
        $rowTemplate->setValue('F', $formula);
    }
    return $rowTemplate;
});

$excel->save($out);

echo 'Output file: ' . $out . '<br>';
echo 'Elapsed time: ' . round(microtime(true) - $time, 3) . ' sec';
