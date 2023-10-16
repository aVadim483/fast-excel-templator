<?php

namespace avadim\FastExcelTemplator;

use avadim\FastExcelHelper\Helper;
use avadim\FastExcelReader\Interfaces\InterfaceSheetReader;
use avadim\FastExcelWriter\Interfaces\InterfaceSheetWriter;

class Sheet extends \avadim\FastExcelReader\Sheet implements InterfaceSheetReader
{
    public SheetWriter $sheetWriter;

    protected array $fill = [];
    protected array $replace = [];
    protected array $rowTemplates = [];
    protected int $insertedRowsCount = 0;
    protected int $lastTouchRowNum = 0;
    protected int $lastReadRowNum = 0;

    protected ?\Generator $readGenerator = null;

    public function __construct($sheetName, $sheetId, $file, $path, $excel)
    {
        parent::__construct($sheetName, $sheetId, $file, $path, $excel);
        $this->preReadFunc = [$this, 'preRead'];
        $this->postReadFunc = [$this, 'postRead'];
    }

    /**
     * @param string|null $file
     *
     * @return Reader
     */
    protected function getReader(string $file = null): Reader
    {
        if (empty($this->xmlReader)) {
            if (!$file) {
                $file = $this->zipFilename;
            }
            $this->xmlReader = Excel::createReader($file);
        }

        return $this->xmlReader;
    }

    public function preRead($xmlReader)
    {
    }

    public function postRead($xmlReader)
    {
        while ($xmlReader->read()) {
            if ($xmlReader->nodeType === \XMLReader::ELEMENT) {
                if ($xmlReader->name === 'mergeCell') {
                    $range = $xmlReader->getAttribute('ref');
                    if ($range) {
                        $this->sheetWriter->mergeCells($range);
                    }
                }
                elseif ($xmlReader->name === 'pageMargins') {
                    $options = $xmlReader->getAllAttributes();
                    if ($options) {
                        $this->sheetWriter->setPageMargins($options);
                    }
                }
                elseif ($xmlReader->name === 'pageSetup') {
                    $options = $xmlReader->getAllAttributes();
                    if ($options) {
                        $this->sheetWriter->setPageSetup($options);
                    }
                }
            }

        }
    }

    /**
     * @param $cell
     * @param $styleIdx
     * @param $formula
     * @param $dataType
     * @param $originalValue
     *
     * @return mixed
     */
    protected function _cellValue($cell, &$styleIdx = null, &$formula = null, &$dataType = null, &$originalValue = null)
    {
        $result = parent::_cellValue($cell, $styleIdx, $formula, $dataType, $originalValue);
        $address = $cell->attributes['r']->value;
        $colIdx = Helper::colNumber($address) - 1;
        $rowIdx = Helper::rowNumber($address) - 1;
        $this->sheetWriter->setNode($rowIdx, $colIdx, $cell);

        return $result;
    }

    /**
     * Replacement for the entire cell value
     *
     * @param array $params
     *
     * @return $this
     */
    public function fill(array $params): Sheet
    {
        $this->sheetWriter->setFillValues($params);

        return $this;
    }

    /**
     * Replacement for substrings in a cell
     *
     * @param array $params
     *
     * @return $this
     */
    public function replace(array $params): Sheet
    {
        $this->sheetWriter->setReplaceValues($params);

        return $this;
    }

    public function getRowTemplate(int $rowNumber): RowTemplate
    {
        if (empty($this->rowTemplates[$rowNumber])) {
            $xmlReader = $this->getReader();
            $xmlReader->openZip($this->path);
            $found = false;
            $rowTemplate = new RowTemplate();
            while ($xmlReader->read()) {
                if ($xmlReader->nodeType === \XMLReader::ELEMENT && $xmlReader->name === 'row' && (int)$xmlReader->getAttribute('r') === $rowNumber) {
                    $found = true;
                    continue;
                }
                if ($xmlReader->nodeType === \XMLReader::END_ELEMENT && $xmlReader->name === 'row' && $found) {
                    break;
                }
                if ($xmlReader->nodeType === \XMLReader::ELEMENT && $xmlReader->name === 'c') {
                    $addr = $xmlReader->getAttribute('r');
                    if ($addr && preg_match('/^([A-Za-z]+)(\d+)$/', $addr, $m)) {
                        $cell = $xmlReader->expand();
                        $value = $this->_cellValue($cell, $styleIdx, $formula, $dataType, $originalValue);
                        $cellData = ['v' => $value, 's' => $styleIdx, 'f' => $formula, 't' => $dataType, 'o' => $originalValue, 'x' => $cell];
                        $rowTemplate->addCell($m[1], $cellData);
                    }
                }
            }
            $this->rowTemplates[$rowNumber] = $rowTemplate;
        }
        $this->lastTouchRowNum = $rowNumber;

        return clone $this->rowTemplates[$rowNumber];
    }

    /**
     * @param int $rowNumber
     * @param mixed $row
     * @param array|null $cellData
     */
    public function insertRow(int $rowNumber, $row, ?array $cellData = [])
    {
        $this->transferRows($rowNumber - 1 - $this->insertedRowsCount);
        if (is_array($row)) {
            $cellData = $row;
            $row = new RowTemplate($cellData);
        }
        else {
            $row->setValues($cellData);
        }
        foreach ($row as $colLetter => $cell) {
            $cellAddress = $colLetter . $rowNumber;
            $cellAddressIdx = ['row_idx' => $rowNumber - 1, 'col_idx' => Helper::colIndex($colLetter)];
            if ($cell instanceof \DOMElement) {
                $value = $cell->nodeValue;
                if ($cell->hasAttributes()) {
                    $styleId = $cell->getAttribute('s');
                    if ($styleId !== '') {
                        $this->sheetWriter->_setStyleIdx($cellAddress, (int)$styleId);
                    }
                }
                ///$this->sheetWriter->writeTo($cellAddress, $value);
                $this->sheetWriter->_writeToCellByIdx($cellAddressIdx, $value);
            }
            elseif (is_array($cell)) {
                $this->_writeWithStyle($cellAddress, $cellAddressIdx, $cell);
            }
            else {
                ///$this->sheetWriter->writeTo($cellAddress, null);
                $this->sheetWriter->_writeToCellByIdx($cellAddressIdx, null);
            }
        }
        $this->insertedRowsCount++;
        $this->lastTouchRowNum = $rowNumber;
        $this->sheetWriter->nextRow();
    }

    /**
     * @param mixed $row
     * @param array|null $cellData
     */
    public function insertRowAfterLast($row, ?array $cellData = [])
    {
        $this->insertRow($this->lastTouchRowNum + 1, $row, $cellData);
    }

    /**
     * @param int $rowNumber
     * @param mixed $row
     * @param array|null $cellData
     */
    public function replaceRow(int $rowNumber, $row, ?array $cellData = [])
    {
        $this->insertRow($rowNumber, $row, $cellData);
        $this->transferRows($rowNumber, true);
        $this->insertedRowsCount--;
    }

    /**
     * @param int $rowNumber
     * @param array|null $cellData
     */
    public function cloneRow(int $rowNumber, ?array $cellData = [])
    {
        $row = $this->getRowTemplate($rowNumber);
        $this->insertRow($rowNumber + 1, $row, $cellData);
    }

    /**
     * Returns the next row from template
     *
     * @return \Generator|null
     */
    public function readRow(): ?\Generator
    {
        if (empty($this->readGenerator)) {
            $this->readGenerator = $this->nextRow([], 0, true);
        }
        while ($rowNum = $this->readGenerator->key()) {
            $cellData = $this->readGenerator->current();
            $this->readGenerator->next();
            $this->lastReadRowNum = $rowNum;
            yield $rowNum => $cellData;
        }

        return null;
    }

    private function _writeWithStyle($cellAddress, $cellAddressIdx, $cellData)
    {
        $numberFormatType = null;
        if ($cellData['t'] === 'date' && is_numeric($cellData['o'])) {
            ///$this->sheetWriter->writeTo($cellAddress, $cellData['o']);
            $this->sheetWriter->_writeToCellByIdx($cellAddressIdx, $cellData['o']);
            $numberFormatType = 'n_auto';
        }
        elseif (!empty($cellData['f'])) {
            ///$this->sheetWriter->writeTo($cellAddress, $cellData['f']);
            $this->sheetWriter->_writeToCellByIdx($cellAddressIdx, $cellData['f']);
        }
        else {
            if ($cellData['t'] === 'date') {
                $pattern = $this->excel->getDateFormatPattern($cellData['s']);
                ///$this->sheetWriter->writeTo($cellAddress, $cellData['v'], ['format' => $pattern]);
                $this->sheetWriter->_writeToCellByIdx($cellAddressIdx, $cellData['v'], ['format' => $pattern]);
                $numberFormatType = 'n_date';
            }
            else {
                ///$this->sheetWriter->writeTo($cellAddress, $cellData['v']);
                $this->sheetWriter->_writeToCellByIdx($cellAddressIdx, $cellData['v']);
            }
        }
        $this->sheetWriter->_setStyleIdx($cellAddress, $cellData['s'], $numberFormatType);
    }

    /**
     * Transfers rows from template to output
     *
     * @param int|null $maxRowNum Max row of template
     * @param bool|null $idle
     *
     * @return void
     */
    public function transferRows(?int $maxRowNum = 0, ?bool $idle = false)
    {
        if (!$maxRowNum || $maxRowNum > $this->lastReadRowNum) {
            foreach ($this->readRow() as $rowNum => $rowData) {
                if (!$idle) {
                    foreach ($rowData as $colLetter => $cellData) {
                        $cellAddress = $colLetter . ($rowNum + $this->insertedRowsCount);
                        $cellAddressIdx = ['row_idx' => $rowNum + $this->insertedRowsCount - 1, 'col_idx' => Helper::colIndex($colLetter)];
                        $this->_writeWithStyle($cellAddress, $cellAddressIdx, $cellData);
                    }
                    $this->sheetWriter->nextRow();
                }
                if (!empty($maxRowNum) && !empty($rowNum) && ($rowNum >= $maxRowNum)) {
                    break;
                }
            }
        }
    }
}