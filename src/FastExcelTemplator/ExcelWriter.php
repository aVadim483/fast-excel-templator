<?php

namespace avadim\FastExcelTemplator;

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

    public function replaceSheets($inputFile, $outputFile): bool
    {
        return $this->writer->_replaceSheets($inputFile, $outputFile);
    }
}