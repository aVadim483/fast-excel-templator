<?php

namespace avadim\FastExcelTemplator;

use avadim\FastExcelHelper\Helper;
use avadim\FastExcelReader\Interfaces\InterfaceSheetReader;
use avadim\FastExcelWriter\Interfaces\InterfaceSheetWriter;

class Sheet extends \avadim\FastExcelReader\Sheet implements InterfaceSheetReader
{
    public SheetWriter $sheetWriter;

    public int $lastReadRowNum = 0;

    public int $countInsertedRows = 0;

    protected array $fill = [];
    protected array $replace = [];
    protected array $rowTemplates = [];
    protected int $lastTouchRowNum = 0;

    protected ?\Generator $readGenerator = null;

    protected array $tables = [];


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
        $tags = ['pageMargins', 'pageSetup', 'drawing', 'legacyDrawing'];
        while ($xmlReader->read()) {
            if ($xmlReader->nodeType === \XMLReader::ELEMENT) {
                if ($xmlReader->name === 'mergeCell') {
                    $range = $xmlReader->getAttribute('ref');
                    if ($range) {
                        $this->sheetWriter->mergeCells($range, 2);
                    }
                }
                elseif (in_array($xmlReader->name, $tags)) {
                    $options = $xmlReader->getAllAttributes();
                    if ($options) {
                        $this->sheetWriter->setBottomNodesOptions($xmlReader->name, $options);
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

    /**
     * @param int $rowNumber
     *
     * @return RowTemplate
     */
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
                    $rowTemplate->setAttributes($xmlReader->getAllAttributes());
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
                        $cellData['__merged'] = $this->mergedRange($addr);
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
        $this->transferRows($rowNumber - 1 - $this->countInsertedRows);
        if (is_array($row)) {
            $cellData = $row;
            $row = new RowTemplate($cellData);
        }
        elseif ($cellData) {
            $row->setValues($cellData);
        }

        $rowHeight = ($row instanceof RowTemplate) ? $row->attribute('ht') : null;
        if ($rowHeight !== null) {
            $this->sheetWriter->setRowHeight($rowNumber, $rowHeight);
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
        $this->countInsertedRows++;
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
        $this->countInsertedRows--;
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
            $this->readGenerator = $this->nextRow([], \avadim\FastExcelReader\Excel::RESULT_MODE_ROW, true);
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
        if (isset($cellData['s'])) {
            $this->sheetWriter->_setStyleIdx($cellAddress, $cellData['s'], $numberFormatType);
        }
        if (isset($cellData['__merged']) && !Helper::inRange($cellAddress, $cellData['__merged'])) {
            $oldRange = $cellData['__merged'];
            $newRange = Helper::addToRange($cellAddress, $cellData['__merged']);
            $this->sheetWriter->updateMergedCells($oldRange, $newRange);
        }
    }

    /**
     * Transfers rows from template to output
     *
     * @param int|null $maxRowNum Max row of template
     * @param bool|null $idle
     *
     * @return void
     */
    public function transferRows(?int $maxRowNum = null, ?bool $idle = false)
    {
        if ($maxRowNum === null || $maxRowNum > $this->lastReadRowNum) {
            foreach ($this->readRow() as $rowNum => $rowData) {
                if (!$idle && (!$maxRowNum || $rowNum <= $maxRowNum)) {
                    $rowNumOut = $rowNum + $this->countInsertedRows;
                    if (isset($rowData['__row']['ht'])) {
                        $this->sheetWriter->setRowHeight($rowNumOut, $rowData['__row']['ht']);
                    }
                    foreach ($rowData['__cells'] as $colLetter => $cellData) {
                        $cellAddress = $colLetter . $rowNumOut;
                        $cellAddressIdx = ['row_idx' => $rowNumOut - 1, 'col_idx' => Helper::colIndex($colLetter)];
                        $this->_writeWithStyle($cellAddress, $cellAddressIdx, $cellData);
                    }
                    $this->sheetWriter->nextRow();
                }
                if ($maxRowNum !== null && !empty($rowNum) && ($rowNum >= $maxRowNum)) {
                    break;
                }
            }
        }
    }

    /**
     * @param string $range
     * @param string|null $header
     * @param string|null $footer
     *
     * @return TableTemplate
     */
    public function table(string $range, ?string $header = null, ?string $footer = null): TableTemplate
    {
        $table = new TableTemplate($this, $range, $header, $footer);
        $this->tables[] = $table;

        return $table;
    }
}
