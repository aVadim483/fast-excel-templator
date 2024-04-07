<?php

namespace avadim\FastExcelTemplator;

use avadim\FastExcelWriter\Exceptions\ExceptionFile;

class ExcelWriter extends \avadim\FastExcelWriter\Excel
{
    public static function createSheet(string $sheetName): SheetWriter
    {
        return new SheetWriter($sheetName);
    }

    /**
     * @param string|null $sheetName
     *
     * @return \avadim\FastExcelWriter\Sheet|SheetWriter
     */
    public function makeSheet(string $sheetName = null): SheetWriter
    {
        return parent::makeSheet($sheetName);
    }

    /**
     * @param $inputFile
     * @param $outputFile
     *
     * @return bool
     */
    public function replaceSheetsAndSave($inputFile, $outputFile): bool
    {
        $relationShips = [
            'rel_id' => ['workbook' => 0],
        ];

        $result = $this->writer->_replaceSheets($inputFile, $outputFile, $relationShips);
        if ($result) {
            $zip = new \ZipArchive();
            if (!$zip->open($outputFile)) {
                ExceptionFile::throwNew('Unable to open zip "%s"', $outputFile);
            }

            if ($zip->statName('xl/calcChain.xml')) {
                $zip->deleteName('xl/calcChain.xml');
                $str = $zip->getFromName('[Content_Types].xml');
                if ($str) {
                    $str = preg_replace('#<[^>]+calcChain.xml[^>]+>#i', '', $str);
                    $zip->addFromString('[Content_Types].xml', $str);
                }
                $str = $zip->getFromName('xl/_rels/workbook.xml.res');
                if ($str) {
                    $str = preg_replace('#<[^>]+calcChain.xml[^>]+>#i', '', $str);
                    $zip->addFromString('xl/_rels/workbook.xml.res', $str);
                }
            }

            $zip->close();
            $this->saved = true;
        }

        return $result;
    }
}