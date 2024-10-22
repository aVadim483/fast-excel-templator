<?php

namespace avadim\FastExcelTemplator;

use avadim\FastExcelHelper\Helper;
use avadim\FastExcelReader\Interfaces\InterfaceSheetReader;
use avadim\FastExcelWriter\Interfaces\InterfaceSheetWriter;

/**
 *
 */
class Sheet extends \avadim\FastExcelReader\Sheet implements InterfaceSheetReader
{
    public SheetWriter $sheetWriter;

    public int $lastReadRowNum = 0;

    public int $lastWrittenRowNum = 0;

    public int $countInsertedRows = 0;

    protected int $topRowOffset;
    protected array $fillValues = [];
    protected array $replaceValues = [];
    protected array $rowTemplates = [];
    protected int $lastTouchRowNum = 0;

    protected ?\Generator $readGenerator = null;

    protected array $sortedMergedCells = [];

    protected ?Reader $rowTemplateReader = null;
    protected int $rowTemplateNo = 0;


    /**
     * @param $sheetName
     * @param $sheetId
     * @param $file
     * @param $path
     * @param $excel
     */
    public function __construct($sheetName, $sheetId, $file, $path, $excel)
    {
        parent::__construct($sheetName, $sheetId, $file, $path, $excel);
        $this->preReadFunc = [$this, 'preRead'];
        $this->postReadFunc = [$this, 'postRead'];
        // init dimension array
        $this->dimension();
        $this->topRowOffset = $this->dimension['min_row_num'] - 1;
        foreach ($this->getMergedCells() as $cell => $range) {
            $cellArr = Helper::rangeArray($cell);
            $this->sortedMergedCells[$cellArr['min_row_num']][$cell] = $range;
        }
        if ($this->sortedMergedCells) {
            ksort($this->sortedMergedCells);
        }
    }

    /**
     * @return string
     */
    public function path(): string
    {
        return $this->pathInZip;
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
        $tags = ['autoFilter', 'pageMargins', 'pageSetup', 'drawing', 'legacyDrawing'];
        //$tags = ['pageMargins', 'pageSetup', 'drawing', 'legacyDrawing'];
        while ($xmlReader->read()) {
            if ($xmlReader->nodeType === \XMLReader::ELEMENT) {
                if (in_array($xmlReader->name, $tags)) {
                    if ($xmlReader->name === 'autoFilter') {
                        $ref = $xmlReader->getAttribute('ref');
                        if ($ref) {
                            //$range = Helper::rangeArray($ref);
                            $this->sheetWriter->setAutofilter(1, 1);
                        }
                    }
                    else {
                        $options = $xmlReader->getAllAttributes();
                        if ($options) {
                            $this->sheetWriter->setBottomNodesOptions($xmlReader->name, $options);
                        }
                    }
                }
            }

        }
    }

    /**
     * @param $cell
     * @param array|null $additionalData
     *
     * @return mixed
     */
    protected function _cellValue($cell, &$additionalData = [])
    {
        $result = parent::_cellValue($cell, $additionalData);
        $address = $cell->attributes['r']->value;
        $colIdx = Helper::colNumber($address) - 1;
        $rowIdx = Helper::rowNumber($address) - 1;
        $this->sheetWriter->setNode($rowIdx, $colIdx, $cell);

        return $result;
    }

    /**
     * Convert A1 addresses to RC in formula
     *
     * @param $node
     * @param string $address
     *
     * @return string
     */
    protected function _cellFormula($node, string $address): string
    {
        $formula = parent::_cellFormula($node, $address);
        $ref = (string)$node->getAttribute('ref');
        if (!$ref) {
            $ref = $address;
        }
        $tokens = token_get_all('<?' . $formula . '?>');
        $tokens[0] = '=';
        $max = count($tokens) - 1;
        unset($tokens[$max]);
        $max--;
        $result = '';
        foreach ($tokens as $n => $t) {
            if (isset($t[0]) && $t[0] === T_STRING && ($n === $max || $tokens[$n + 1] !== '(') && preg_match('/^[A-Z]+[0-9]+$/', $t[1])) {
                $result .= Helper::A1toRC($t[1], $ref);
            }
            else {
                $result .= (is_string($t) ? $t : $t[1]);
            }
        }

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
     * Returns the rel range of merged cells that contains the specified cell
     *
     * @param string $cellAddress
     *
     * @return string|null
     */
    public function mergedRange(string $cellAddress): ?string
    {
        $result = parent::mergedRange($cellAddress);
        if ($result && strpos($result, $cellAddress . ':') === 0) {
            $adr = Helper::rangeArray($cellAddress);
            $dim = Helper::rangeArray($result);
            $rowOffset1 = $dim['min_row_num'] - $adr['min_row_num'];
            $colOffset1 = $dim['min_col_num'] - $adr['min_col_num'];
            $rowOffset2 = $dim['max_row_num'] - $adr['min_row_num'];
            $colOffset2 = $dim['max_col_num'] - $adr['min_col_num'];
            $result = 'R' . (($rowOffset1 >= 0) ? $rowOffset1 : '[' . $rowOffset1 . ']')
                . 'C' . (($colOffset1 >= 0) ? $colOffset1 : '[' . $colOffset1 . ']') . ':'
                . 'R' . (($rowOffset2 >= 0) ? $rowOffset2 : '[' . $rowOffset2 . ']')
                . 'C' . (($colOffset2 >= 0) ? $colOffset2 : '[' . $colOffset2 . ']');

        }
        else {
            $result = null;
        }

        return $result;
    }

    /**
     * @param int $rowNumberMin
     * @param int $rowNumberMax
     *
     * @return Reader
     */
    protected function getRowTemplateReader(int $rowNumberMin, int $rowNumberMax): Reader
    {
        if ($rowNumberMax < $rowNumberMin) {
            throw new \RuntimeException('$rowNumberMax cannot be less then $rowNumberMin');
        }
        if ($rowNumberMin < $this->dimension['min_row_num']) {
            throw new \RuntimeException('$rowNumberMin cannot be less then ' . $this->dimension['min_row_num']);
        }
        if ($rowNumberMin > $this->dimension['max_row_num']) {
            throw new \RuntimeException('$rowNumberMin cannot be more then ' . $this->dimension['max_row_num']);
        }
        if ($rowNumberMax < $this->dimension['min_row_num']) {
            throw new \RuntimeException('$rowNumberMax cannot be less then ' . $this->dimension['min_row_num']);
        }
        if ($rowNumberMax > $this->dimension['max_row_num']) {
            throw new \RuntimeException('$rowNumberMax cannot be more then ' . $this->dimension['max_row_num']);
        }

        if (!empty($this->rowTemplateReader) && $this->rowTemplateNo > $rowNumberMin && !isset($this->rowTemplates[$rowNumberMin])) {
            // Need to reset reader
            $this->rowTemplateReader = null;
        }
        if (empty($this->rowTemplateReader)) {
            $this->rowTemplateReader = Excel::createReader($this->zipFilename);
            $this->rowTemplateReader->openZip($this->path());
        }

        return $this->rowTemplateReader;
    }

    /**
     * @param int $rowNumber
     * @param bool|null $savePointerPosition
     *
     * @return RowTemplateCollection
     */
    public function getRowTemplate(int $rowNumber, ?bool $savePointerPosition = false): RowTemplateCollection
    {
        return $this->getRowTemplates($rowNumber, $rowNumber, $savePointerPosition);
    }

    /**
     * @param int $rowNumberMin
     * @param int $rowNumberMax
     * @param bool|null $savePointerPosition
     *
     * @return RowTemplateCollection
     */
    public function getRowTemplates(int $rowNumberMin, int $rowNumberMax, ?bool $savePointerPosition = false): RowTemplateCollection
    {
        $findNum = [];
        for ($rowNum = $rowNumberMin; $rowNum <= $rowNumberMax; $rowNum++) {
            if (!isset($this->rowTemplates[$rowNum])) {
                $findNum[$rowNum] = $rowNum;
            }
        }
        if ($findNum) {
            $xmlReader = $this->getRowTemplateReader($rowNumberMin, $rowNumberMax);

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
                                    $this->_cellValue($cell, $additionalData);
                                    $cellData = $additionalData;
                                    $cellData['__address'] = $addr;
                                    $cellData['__merged'] = $this->mergedRange($addr);
                                    $rowTemplate->addCell($m[1], $cellData);
                                }
                            }
                        }
                        unset($findNum[$r]);
                        $this->rowTemplates[$r] = $rowTemplate;
                        $this->rowTemplateNo = $r;
                    }
                }
                if (!$findNum) {
                    break;
                }
            }
        }

        $rows = [];
        for ($rowNum = $rowNumberMin; $rowNum <= $rowNumberMax; $rowNum++) {
            $rows[$rowNum] = clone $this->rowTemplates[$rowNum];
        }

        $this->lastTouchRowNum = $rowNumberMax;
        if (!$savePointerPosition) {
            $this->skipRowsUntil($rowNumberMax);
        }

        $rowsCollection = new RowTemplateCollection($rows);
        $rowsCollection->setSheet($this);

        return $rowsCollection;
    }

    /**
     * @param array|RowTemplateCollection|RowTemplate $row
     * @param array|null $cellData
     *
     * @return Sheet
     */
    public function insertRow($row, ?array $cellData = []): Sheet
    {
        if (is_array($row)) {
            $cellData = $row;
            $row = new RowTemplate();
        }
        elseif ($row instanceof RowTemplateCollection) {
            $row = $row->next();
        }
        $row->setValues($cellData);

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
                $this->sheetWriter->_writeToCellByIdx($cellAddressIdx, $value);
            }
            elseif (is_array($cell)) {
                $this->_writeWithStyle($cellAddress, $cellAddressIdx, $cell);
            }
            else {
                $this->sheetWriter->_writeToCellByIdx($cellAddressIdx, null);
            }
        }
        $this->countInsertedRows++;
        $this->lastTouchRowNum = $rowNumber;
        $this->sheetWriter->nextRow();

        return $this;
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
            $rowData = $this->readGenerator->current();
            $rowTemplate = new RowTemplate();
            if (isset($rowData['__row'])) {
                $rowTemplate->setAttributes($rowData['__row']);
            }
            foreach ($rowData['__cells'] as $col => $cellData) {
                $sourceAddress = $col . $rowNum;
                $cellData['__sourceAddress'] = $sourceAddress;
                if (isset($this->sortedMergedCells[$rowNum]) && ($merged = $this->mergedRange($sourceAddress))) {
                    $cellData['__merged'] = $merged;
                }
                $rowTemplate->addCell($col, $cellData);
            }
            $this->readGenerator->next();
            $this->lastReadRowNum = $rowNum;
            yield $rowNum => $rowTemplate;
        }

        return null;
    }


    /**
     * Write values to the current row
     *
     * @param array $rowValues Values of cells
     * @param array|null $rowStyle Style applied to the entire row
     * @param array|null $cellStyles Styles of specified cells in the row
     *
     * @return $this
     */
    public function writeRow(array $rowValues = [], array $rowStyle = null, array $cellStyles = null): Sheet
    {
        $this->sheetWriter->writeRow($rowValues, $rowStyle, $cellStyles);

        return $this;
    }

    /**
     * Last read row number
     *
     * @return int
     */
    public function lastReadRowNum(): int
    {
        return $this->lastReadRowNum;
    }

    public function lastWrittenRowNum(): int
    {
        return $this->sheetWriter->currentRowNum() - 1;
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
        if (!empty($cellData['f'])) {
            $this->sheetWriter->_writeToCellByIdx($cellAddressIdx, $cellData['f']);
        }
        else {
            if ($cellData['t'] === 'date') {
                $pattern = $this->excel->getDateFormatPattern($cellData['s']);
                $this->sheetWriter->_writeToCellByIdx($cellAddressIdx, $cellData['v'], ['format' => $pattern]);
                $numberFormatType = 'n_date';
            }
            else {
                $this->sheetWriter->_writeToCellByIdx($cellAddressIdx, $cellData['v']);
            }
        }
        if (isset($cellData['s'])) {
            $this->sheetWriter->_setStyleIdx($cellAddress, $cellData['s'], $numberFormatType);
        }
        if (isset($cellData['__merged'])) {
            $mergedRange = Helper::addToRange($cellAddress, $cellData['__merged']);
            $this->sheetWriter->mergeCells($mergedRange, 1);
        }
    }

    /**
     * Transfers rows from template to output
     *
     * @param int|null $maxRowNum Max row of template
     * @param $callback
     *
     * @return Sheet
     */
    public function transferRowsUntil(?int $maxRowNum = null, $callback = null): Sheet
    {
        if ($maxRowNum === null || $maxRowNum > $this->lastReadRowNum) {
            foreach ($this->readRow() as $sourceRowNum => $rowData) {
                if (!$maxRowNum || $sourceRowNum <= $maxRowNum) {
                    $targetRowNum = $this->sheetWriter->currentRowNum();
                    if ($targetRowNum < $sourceRowNum) {
                        $targetRowNum = $sourceRowNum;
                    }
                    if ($callback) {
                        $rowData = $callback($targetRowNum, $rowData);
                    }
                    if ($height = $rowData->rowHeight()) {
                        $this->sheetWriter->setRowHeight($targetRowNum, $height);
                    }
                    foreach ($rowData->cells() as $colLetter => $cellData) {
                        $cellAddress = $colLetter . $targetRowNum;
                        $cellAddressIdx = ['row_idx' => $targetRowNum - 1, 'col_idx' => Helper::colIndex($colLetter)];
                        $this->_writeWithStyle($cellAddress, $cellAddressIdx, $cellData);
                    }
                    $this->sheetWriter->nextRow();
                }
                if ($maxRowNum !== null && !empty($sourceRowNum) && ($sourceRowNum >= $maxRowNum)) {
                    while ($this->sheetWriter->currentRowNum() <= $maxRowNum) {
                        $this->sheetWriter->nextRow();
                    }
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
     * @param $callback
     *
     * @return Sheet
     */
    public function transferRows(?int $countRows = null, $callback = null): Sheet
    {
        return $this->transferRowsUntil($countRows ? ($this->lastReadRowNum + $countRows) : null, $callback);
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


    public function saveSheet()
    {
    }
}
