<?php

namespace avadim\FastExcelTemplator;

use avadim\FastExcelHelper\Helper;
use avadim\FastExcelReader\Interfaces\InterfaceSheetReader;
use avadim\FastExcelWriter\Interfaces\InterfaceSheetWriter;

class Sheet extends \avadim\FastExcelReader\Sheet implements InterfaceSheetReader
{
    public SheetWriter $sheetWriter;

    public int $lastReadRowNum = 0;

    public int $lastWrittenRowNum = 0;

    public int $countInsertedRows = 0;

    protected int $rowOffset;
    protected array $fillValues = [];
    protected array $replaceValues = [];
    protected array $rowTemplates = [];
    protected int $lastTouchRowNum = 0;

    protected ?\Generator $readGenerator = null;

    protected array $tables = [];


    public function __construct($sheetName, $sheetId, $file, $path, $excel)
    {
        parent::__construct($sheetName, $sheetId, $file, $path, $excel);
        $this->preReadFunc = [$this, 'preRead'];
        $this->postReadFunc = [$this, 'postRead'];
        // init dimension array
        $this->dimension();
        $this->rowOffset = $this->dimension['min_row_num'] - 1;
    }

    /**
     * @return string
     */
    public function path(): string
    {
        return $this->path;
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
    public function fillValues(array $params): Sheet
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
    public function replaceValues(array $params): Sheet
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
        $rowTemplates = $this->getRowTemplates($rowNumber, $rowNumber);

        return $rowTemplates[$rowNumber];
    }

    /**
     * @param int $rowNumberMin
     * @param int $rowNumberMax
     *
     * @return array
     */
    public function getRowTemplates(int $rowNumberMin, int $rowNumberMax): array
    {
        $findNum = [];
        for ($rowNum = $rowNumberMin; $rowNum <= $rowNumberMax; $rowNum++) {
            if (!isset($this->rowTemplates[$rowNum])) {
                $findNum[$rowNum] = $rowNum;
            }
        }
        if ($findNum) {
            $xmlReader = Excel::createReader($this->zipFilename);
            $xmlReader->openZip($this->path);

            while ($xmlReader->read()) {
                if ($xmlReader->nodeType === \XMLReader::ELEMENT && $xmlReader->name === 'row') {
                    $r = (int)$xmlReader->getAttribute('r');
                    if (isset($findNum[$r])) {
                        $rowTemplate = new RowTemplate();
                        $rowTemplate->setAttributes($xmlReader->getAllAttributes());
                        while ($xmlReader->read() && !($xmlReader->nodeType === \XMLReader::END_ELEMENT && $xmlReader->name === 'row')) {
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
                        unset($findNum[$r]);
                        $this->rowTemplates[$r] = $rowTemplate;
                    }
                }
                if (!$findNum) {
                    break;
                }
            }
        }
        $result = [];
        for ($rowNum = $rowNumberMin; $rowNum <= $rowNumberMax; $rowNum++) {
            $result[$rowNum] = clone $this->rowTemplates[$rowNum];
        }

        $this->lastTouchRowNum = $rowNumberMax;

        return $result;
    }

    /**
     * @param mixed $row
     * @param array|null $cellData
     */
    public function insertRow($row, ?array $cellData = [])
    {
        if (is_array($row)) {
            $cellData = $row;
            $row = new RowTemplate($cellData);
        }
        elseif ($cellData) {
            $row->setValues($cellData);
        }
        $rowNumber = $this->sheetWriter->currentRowNum();
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
    public function replaceRow($row, ?array $cellData = [])
    {
        $this->insertRow($row, $cellData);
        foreach ($this->readRow() as $rowNum => $rowData) {
            break;
        }
    }

    /**
     * @param array|null $cellData
     */
    public function cloneRow(?array $cellData = [])
    {
        $row = $this->getRowTemplate($this->lastReadRowNum);
        $this->insertRow($row, $cellData);
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

    /**
     * Last read row number
     *
     * @return int
     */
    public function latReadRowNum(): int
    {
        return $this->lastReadRowNum;
    }

    /**
     * Write cell value and style
     *
     * @param $cellAddress
     * @param $cellAddressIdx
     * @param $cellData
     *
     * @return void
     */
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
     *
     * @return Sheet
     */
    public function transferRowsUntil(?int $maxRowNum = null): Sheet
    {
        if ($maxRowNum === null || $maxRowNum > $this->lastReadRowNum) {
            foreach ($this->readRow() as $rowNum => $rowData) {
                if (!$maxRowNum || $rowNum <= $maxRowNum) {
                    $rowNumOut = $this->sheetWriter->currentRowNum();
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
        return $this;
    }

    /**
     * Transfers rows from template to output
     *
     * @param int|null $countRows Number of rows
     *
     * @return Sheet
     */
    public function transferRows(?int $countRows = null): Sheet
    {
        return $this->transferRowsUntil($countRows ? ($this->lastReadRowNum + $countRows) : null);
    }

    /**
     * Skip rows from template
     *
     * @param int|null $maxRowNum Max row of template
     *
     * @return Sheet
     */
    public function skipRowsUntil(?int $maxRowNum = null): Sheet
    {
        if ($maxRowNum === null || $maxRowNum > $this->lastReadRowNum) {
            foreach ($this->readRow() as $rowNum => $rowData) {
                if ($maxRowNum !== null && !empty($rowNum) && ($rowNum >= $maxRowNum)) {
                    break;
                }
            }
        }
        return $this;
    }

    /**
     * Skip rows from template
     *
     * @param int|null $countRows Number of rows
     *
     * @return Sheet
     */
    public function skipRows(?int $countRows = null): Sheet
    {
        return $this->skipRowsUntil($countRows ? ($this->lastReadRowNum + $countRows) : null);
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
        $this->tables[$table->tplRange['min_row_num']] = $table;

        return $table;
    }


    public function saveSheet()
    {
        if ($this->tables) {
            ksort($this->tables);
            /**
             * @var int $tableRowBegin
             * @var TableTemplate $table
             */
            foreach ($this->tables as $table) {
                $table->transferRows();
            }
        }
    }
}
