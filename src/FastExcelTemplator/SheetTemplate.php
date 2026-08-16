<?php

namespace avadim\FastExcelTemplator;

use avadim\FastExcelHelper\Helper;
use avadim\FastExcelReader\Interfaces\InterfaceSheetReader;
use avadim\FastExcelWriter\DataValidation\DataValidation;
use avadim\FastExcelWriter\Exceptions\ExceptionDataValidation;
use avadim\FastExcelWriter\Interfaces\InterfaceSheetWriter;

/**
 *
 */
class SheetTemplate extends \avadim\FastExcelReader\Sheet implements InterfaceSheetReader
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
    protected bool $mergedCellsInit = false;

    protected ?Reader $rowTemplateReader = null;
    protected int $rowTemplateNo = 0;

    protected bool $postReadDone  = false;


    /**
     * SheetTemplate constructor
     *
     * @param string $sheetName
     * @param string $sheetId
     * @param string $file
     * @param string $path
     * @param Excel $excel
     */
    public function __construct(string $sheetName, string $sheetId, string $file, string $path, $excel)
    {
        parent::__construct($sheetName, $sheetId, $file, $path, $excel);
        $this->preReadFunc = [$this, 'preRead'];
        $this->postReadFunc = [$this, 'postRead'];
        // init dimension array
        $this->dimension();
        $this->topRowOffset = isset($this->dimension['min_row_num']) ? $this->dimension['min_row_num'] - 1 : 0;
        // NOTE: merged cells are read lazily (see initSortedMergedCells()) to avoid
        // a full scan of every sheet at open time, especially for sheets that are never transferred
    }

    /**
     * Lazily build the map of merged cells sorted by their top row.
     *
     * Reading merged cells requires a full scan of the sheet up to the end of the file
     * (see Reader::_readBottom()), so it is deferred until the merges are actually needed
     * during row reading instead of being done eagerly in the constructor.
     *
     * @return void
     */
    protected function initSortedMergedCells(): void
    {
        if ($this->mergedCellsInit) {
            return;
        }
        $this->mergedCellsInit = true;
        foreach ($this->getMergedCells() as $cell => $range) {
            $cellArr = Helper::rangeArray($cell);
            $this->sortedMergedCells[$cellArr['min_row_num']][$cell] = $range;
        }
        if ($this->sortedMergedCells) {
            ksort($this->sortedMergedCells);
        }
    }

    /**
     * Get path in ZIP
     *
     * @return string
     */
    public function path(): string
    {
        return $this->pathInZip;
    }

    /**
     * Internal reader
     *
     * @param string|null $file
     *
     * @return Reader
     */
    protected function getReader(?string $file = null): Reader
    {
        if (empty($this->xmlReader)) {
            if (!$file) {
                $file = $this->zipFilename;
            }
            $this->xmlReader = Excel::createReader($file);
        }

        return $this->xmlReader;
    }

    /**
     * Internal pre-read handler
     *
     * @param \XMLReader $xmlReader
     */
    public function preRead($xmlReader)
    {
    }

    /**
     * Internal post-read handler
     *
     * @param \XMLReader $xmlReader
     */
    public function postRead($xmlReader)
    {
        if ($this->postReadDone) {
            return;
        }

        // stop on the closing </sheetData>; the loop below reads the following sibling nodes
        // (note: we must NOT skip the first sibling — e.g. <autoFilter> usually comes right here)
        do {
            if ($xmlReader->nodeType === \XMLReader::END_ELEMENT && $xmlReader->name === 'sheetData') {
                break;
            }
        } while ($xmlReader->read());

        $tags = ['autoFilter', 'dataValidations', 'pageMargins', 'pageSetup', 'drawing', 'legacyDrawing', 'headerFooter'];
        while ($xmlReader->read()) {
            if ($xmlReader->nodeType === \XMLReader::ELEMENT) {
                if (in_array($xmlReader->name, $tags)) {
                    if ($xmlReader->name === 'autoFilter') {
                        $ref = $xmlReader->getAttribute('ref');
                        if ($ref) {
                            // keep the original filter range from the template instead of a hardcoded A1
                            $this->sheetWriter->setAutoFilter($ref);
                        }
                    }
                    elseif ($xmlReader->name === 'dataValidations') {
                        $this->_transferDataValidations($xmlReader->expand());
                    }
                    elseif ($xmlReader->name === 'headerFooter') {
                        $options = $xmlReader->getAllAttributes();
                        $node = $xmlReader->expand();
                        foreach ($node->childNodes as $child) {
                            $options[$child->nodeName] = str_replace('&amp;', '&', $child->nodeValue);
                        }
                        $this->sheetWriter->setHeaderFooterOptions($options);
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
        $this->postReadDone = true;
    }

    /**
     * Transfer the template's <dataValidations> (e.g. dropdown lists) to the output sheet.
     *
     * @param \DOMNode|null $node The <dataValidations> element captured from the template
     *
     * @return void
     */
    private function _transferDataValidations($node): void
    {
        if (!$node instanceof \DOMElement) {
            return;
        }
        foreach ($node->childNodes as $dv) {
            if (!$dv instanceof \DOMElement || $dv->nodeName !== 'dataValidation') {
                continue;
            }
            $type = $dv->getAttribute('type');
            $sqref = $dv->getAttribute('sqref');
            // 'type' and 'sqref' are required to rebuild a validation
            if ($type === '' || $sqref === '') {
                continue;
            }
            try {
                $validation = DataValidation::make($type);
                $operator = $dv->getAttribute('operator');
                if ($operator !== '') {
                    $validation->setOperator($operator);
                }
                foreach ($dv->childNodes as $formula) {
                    if ($formula->nodeName === 'formula1') {
                        $validation->setFormula1($formula->nodeValue);
                    }
                    elseif ($formula->nodeName === 'formula2') {
                        $validation->setFormula2($formula->nodeValue);
                    }
                }
                $validation->allowBlank($dv->getAttribute('allowBlank') === '1');
                // sqref may list several ranges separated by spaces
                foreach (explode(' ', $sqref) as $range) {
                    if ($range !== '') {
                        $this->sheetWriter->addDataValidation($range, $validation);
                    }
                }
            }
            catch (ExceptionDataValidation $e) {
                // The template may hold a type, operator or formula this writer cannot express;
                // skip that single validation rather than break the whole save. Only the writer's
                // own validation error is caught here, so a genuine bug still surfaces.
                continue;
            }
        }
    }

    /**
     * @param $cell
     * @param array|null $additionalData
     *
     * @return mixed
     */
    protected function _cellValue($cell, ?array &$additionalData = [])
    {
        $result = parent::_cellValue($cell, $additionalData);
        // The source DOM node is only needed later for error cells (see SheetWriter::_writeToCellByIdx()),
        // so we avoid retaining a DOMElement for every single cell of the sheet
        if (isset($additionalData['t']) && $additionalData['t'] === 'error') {
            $address = $cell->attributes['r']->value;
            $colIdx = Helper::colNumber($address) - 1;
            $rowIdx = Helper::rowNumber($address) - 1;
            $this->sheetWriter->setNode($rowIdx, $colIdx, $cell);
        }

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
        // an array formula is based on the top-left cell of its range, not on the current cell
        $ref = (string)$node->getAttribute('ref');
        if (!$ref) {
            $ref = $address;
        }

        return FormulaLexer::convertA1toRC($formula, $ref);
    }

    /**
     * Replace the entire cell value: applies only when the whole cell equals the key ('{{X}}', not 'text {{X}}').
     *
     * @param array $params Map of [placeholder => value]; applied to every cell written to the output
     *
     * @return $this
     *
     * @example
     * $sheet->fill(['{{COMPANY}}' => 'ACME Inc.', '{{NUMBER}}' => 128]);
     */
    public function fill(array $params): SheetTemplate
    {
        $this->sheetWriter->setFillValues($params);

        return $this;
    }

    /**
     * Replace a substring inside cell values: the key is matched anywhere in the text ('Date: {{DATE}}' works).
     *
     * @param array $params Map of [search => replacement]; applied to every cell written to the output
     *
     * @return $this
     *
     * @example
     * $sheet->replace(['{{DATE}}' => date('d.m.Y')]);
     */
    public function replace(array $params): SheetTemplate
    {
        $this->sheetWriter->setReplaceValues($params);

        return $this;
    }

    /**
     * Offsets of the merged range that STARTS at the specified cell, or null
     *
     * The result is not an address range but a set of relative offsets ('R0C0:R0C2' means
     * "two columns to the right of this very cell"), so that a captured row template can be
     * re-merged at whatever row it is inserted into — see Helper::addToRange() in _writeWithStyle().
     * Only the top-left cell of a merge answers; every other cell of the same merge returns null,
     * because a merge has to be recreated once, from its own corner.
     *
     * Do not confuse it with mergedRange() inherited from the reader: that one answers for any
     * cell inside a merge and returns a real address range ('A2:C2').
     *
     * @param string $cellAddress
     *
     * @return string|null
     */
    protected function mergedRangeOffsets(string $cellAddress): ?string
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
     * Get internal row template reader
     *
     * @param int $rowNumberMin
     * @param int $rowNumberMax
     *
     * @return Reader
     */
    protected function getRowTemplateReader(int $rowNumberMin, int $rowNumberMax): Reader
    {
        if ($rowNumberMax < $rowNumberMin) {
            throw new Exception('$rowNumberMax cannot be less then $rowNumberMin');
        }
        // Skip boundary checks when the sheet has no <dimension> tag (min/max row numbers are unknown)
        if (isset($this->dimension['min_row_num'], $this->dimension['max_row_num'])) {
            if ($rowNumberMin < $this->dimension['min_row_num']) {
                throw new Exception('$rowNumberMin cannot be less then ' . $this->dimension['min_row_num']);
            }
            if ($rowNumberMin > $this->dimension['max_row_num']) {
                throw new Exception('$rowNumberMin cannot be more then ' . $this->dimension['max_row_num']);
            }
            if ($rowNumberMax < $this->dimension['min_row_num']) {
                throw new Exception('$rowNumberMax cannot be less then ' . $this->dimension['min_row_num']);
            }
            if ($rowNumberMax > $this->dimension['max_row_num']) {
                throw new Exception('$rowNumberMax cannot be more then ' . $this->dimension['max_row_num']);
            }
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
     * Capture a source row as a reusable template; returns a collection (use it with insertRow()).
     *
     * @param int $rowNumber Source row number to capture
     * @param bool|null $savePointerPosition If false (default), the read cursor is advanced to this row
     *
     * @return RowTemplateCollection
     *
     * @example
     * $tpl = $sheet->getRowTemplate(7);
     * foreach ($data as $row) { $sheet->insertRow($tpl, ['A' => $row['id']]); }
     */
    public function getRowTemplate(int $rowNumber, ?bool $savePointerPosition = false): RowTemplateCollection
    {
        return $this->getRowTemplates($rowNumber, $rowNumber, $savePointerPosition);
    }

    /**
     * Capture a range of source rows as templates; insertRow() cycles through them (e.g. for zebra striping).
     *
     * @param int $rowNumberMin First source row of the range
     * @param int $rowNumberMax Last source row of the range
     * @param bool|null $savePointerPosition If false (default), the read cursor is advanced to $rowNumberMax
     *
     * @return RowTemplateCollection
     *
     * @example
     * $tpls = $sheet->getRowTemplates(7, 8); // two styles, alternated on each insertRow()
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
            // same guard as in readRow(): consult the merge map only for rows where a merge starts,
            // instead of scanning every merge of the sheet for every single cell
            $this->initSortedMergedCells();
            $xmlReader = $this->getRowTemplateReader($rowNumberMin, $rowNumberMax);

            while ($xmlReader->read()) {
                if ($xmlReader->nodeType === \XMLReader::ELEMENT && $xmlReader->name === 'row') {
                    $r = (int)$xmlReader->getAttribute('r');
                    if (isset($findNum[$r])) {
                        $rowTemplate = (new RowTemplate())->setSheetTemplate($this);
                        $rowTemplate->setAttributes($xmlReader->getAllAttributes());
                        while ($xmlReader->read() && !($xmlReader->nodeType === \XMLReader::END_ELEMENT && $xmlReader->name === 'row')) {
                            if ($xmlReader->nodeType === \XMLReader::ELEMENT && $xmlReader->name === 'c') {
                                $addr = $xmlReader->getAttribute('r');
                                if ($addr && preg_match('/^([A-Za-z]+)(\d+)$/', $addr, $m)) {
                                    $cell = $xmlReader->expand();
                                    $this->_cellValue($cell, $additionalData);
                                    $cellData = $additionalData;
                                    $cellData['__address'] = $addr;
                                    $cellData['__merged'] = isset($this->sortedMergedCells[$r])
                                        ? $this->mergedRangeOffsets($addr)
                                        : null;
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

        return new RowTemplateCollection($rows, $this);
    }

    /**
     * Insert a row at the current output position, filling it with data by column letter.
     *
     * @param array|RowTemplateCollection|RowTemplate $row A template (or collection to cycle through), or a plain [col => value] array for a blank row
     * @param array|null $cellData Values keyed by column letter, e.g. ['A' => 1, 'B' => 'name']
     *
     * @return SheetTemplate
     *
     * @example
     * $sheet->insertRow($rowTemplate, ['A' => $item['id'], 'B' => $item['name']]);
     */
    public function insertRow($row, ?array $cellData = []): SheetTemplate
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
     * Replace row with data
     *
     * @param mixed $row
     * @param array|null $cellData
     *
     * @return $this
     */
    public function replaceRow($row, ?array $cellData = [])
    {
        $this->insertRow($row, $cellData);
        // consume exactly one source row, so the replacement takes its place
        foreach ($this->readRow() as $rowNum => $rowData) {
            break;
        }

        return $this;
    }

    /**
     * Clone current row
     *
     * @param array|null $cellData
     *
     * @return $this
     */
    public function cloneRow(?array $cellData = [])
    {
        $row = $this->getRowTemplate($this->lastReadRowNum);
        $this->insertRow($row, $cellData);

        return $this;
    }

    /**
     * Returns the next row from template
     *
     * @return \Generator|null
     */
    public function readRow(): ?\Generator
    {
        $this->initSortedMergedCells();
        if (empty($this->readGenerator)) {
            $this->readGenerator = $this->nextRow([], \avadim\FastExcelReader\Excel::RESULT_MODE_ROW, true);
        }
        while ($rowNum = $this->readGenerator->key()) {
            $rowData = $this->readGenerator->current();
            $rowTemplate = (new RowTemplate())->setSheetTemplate($this);
            if (isset($rowData['__row'])) {
                $rowTemplate->setAttributes($rowData['__row']);
            }
            foreach ($rowData['__cells'] as $col => $cellData) {
                $sourceAddress = $col . $rowNum;
                $cellData['__sourceAddress'] = $sourceAddress;
                if (isset($this->sortedMergedCells[$rowNum]) && ($merged = $this->mergedRangeOffsets($sourceAddress))) {
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
     * Last read row number
     *
     * @return int
     */
    public function lastReadRowNum(): int
    {
        return $this->lastReadRowNum;
    }

    /**
     * Last written row number
     *
     * @return int
     */
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
            // 't' may be missing in a cell built by hand (e.g. RowTemplate::addCell() with an array)
            $cellType = $cellData['t'] ?? null;
            $cellValue = $cellData['v'] ?? null;
            if ($cellType === 'date') {
                $pattern = $this->excel->getDateFormatPattern($cellData['s'] ?? 0);
                $this->sheetWriter->_writeToCellByIdx($cellAddressIdx, $cellValue, ['format' => $pattern]);
                $numberFormatType = 'n_date';
            }
            elseif ($cellType === 'error' && $cellValue && $cellValue[0] === '#') {
                $this->sheetWriter->_writeToCellByIdx($cellAddressIdx, ['v' => $cellValue, 't' => 'e']);
                if (empty($cellData['s'])) {
                    $cellData['s'] = 0;
                }
                $numberFormatType = 'n_error';
            }
            else {
                $this->sheetWriter->_writeToCellByIdx($cellAddressIdx, $cellValue);
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
     * Copy source rows to the output up to $maxRowNum; an optional callback can modify, skip or stop each row.
     *
     * @param int|null $maxRowNum Last source row to transfer; null transfers to the end
     * @param callable|null $callback function($sourceRowNum, $targetRowNum, RowTemplate $row): return $row to write it, null to skip the row, false to stop
     *
     * @return SheetTemplate
     */
    public function transferRowsUntil(?int $maxRowNum = null, $callback = null): SheetTemplate
    {
        $skippedRows = 0;
        if ($maxRowNum === null || $maxRowNum > $this->lastReadRowNum) {
            foreach ($this->readRow() as $sourceRowNum => $rowData) {
                if (!$maxRowNum || $sourceRowNum <= $maxRowNum) {
                    $targetRowNum = $this->sheetWriter->currentRowNum();
                    if ($targetRowNum < $sourceRowNum) {
                        $targetRowNum = $sourceRowNum;
                    }
                    if ($callback) {
                        $res = $callback($sourceRowNum, $targetRowNum - $skippedRows, $rowData);
                        if ($res === null) {
                            $skippedRows++;
                            continue;
                        }
                        if ($res === false) {
                            break;
                        }
                        if ($res instanceof RowTemplate) {
                            $rowData = $res;
                        }
                    }
                    $this->sheetWriter->setRowAttributes($targetRowNum - $skippedRows, $rowData->getAttributes());
                    if ($height = $rowData->rowHeight()) {
                        $this->sheetWriter->setRowHeight($targetRowNum - $skippedRows, $height);
                    }
                    $offsetColIdx = 0;
                    foreach ($rowData->cells() as $colLetter => $cellData) {
                        if (!empty($cellData['__removed'])) {
                            $offsetColIdx++;
                        }
                        else {
                            $cellAddress = $colLetter . ($targetRowNum - $skippedRows);
                            $cellAddressIdx = ['row_idx' => $targetRowNum - $skippedRows - 1, 'col_idx' => Helper::colIndex($colLetter) - $offsetColIdx];
                            $this->_writeWithStyle($cellAddress, $cellAddressIdx, $cellData);
                        }
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

            ///$xmlReader = $this->getReader();
            // read page options and others
            ///$this->postRead($xmlReader);
        }
        return $this;
    }

    /**
     * Copy the next $countRows source rows (or all remaining if null) to the output, with an optional callback.
     *
     * @param int|null $countRows Number of rows to transfer; null transfers all remaining rows
     * @param callable|null $callback function($sourceRowNum, $targetRowNum, RowTemplate $row): return $row to write it, null to skip the row, false to stop
     *
     * @return SheetTemplate
     */
    public function transferRows(?int $countRows = null, $callback = null): SheetTemplate
    {
        return $this->transferRowsUntil($countRows ? ($this->lastReadRowNum + $countRows) : null, $callback);
    }

    /**
     * Walk all source rows through a callback and write the result (read-modify-write over the whole sheet).
     *
     * @param callable $callback function($sourceRowNum, $targetRowNum, RowTemplate $row): return $row to write it, null to skip the row, false to stop
     *
     * @return SheetTemplate
     *
     * @example
     * $sheet->rows(function ($src, $dst, $row) {
     *     if ($src === 1) return $row;            // keep header
     *     $row->setValue('C', $row->getValue('C') * 1.2);
     *     return $row;
     * });
     */
    public function rows($callback): SheetTemplate
    {
        return $this->transferRows(null, $callback);
    }

    /**
     * Skip rows from the template
     *
     * @param int|null $maxRowNum Max row of template
     *
     * @return SheetTemplate
     */
    public function skipRowsUntil(?int $maxRowNum = null): SheetTemplate
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
     * Skip rows from the template
     *
     * @param int|null $countRows Number of rows
     *
     * @return SheetTemplate
     */
    public function skipRows(?int $countRows = null): SheetTemplate
    {
        return $this->skipRowsUntil($countRows ? ($this->lastReadRowNum + $countRows) : null);
    }

    /**
     * Empty method for compatibility
     *
     * @return void
     */
    public function saveSheet()
    {

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
    public function writeRow(array $rowValues = [], ?array $rowStyle = null, ?array $cellStyles = null): SheetTemplate
    {
        $this->sheetWriter->writeRow($rowValues, $rowStyle, $cellStyles);

        return $this;
    }

    /**
     * Write value to the current cell and move a pointer to the next cell in the row
     *
     * @param mixed $value
     * @param array|null $styles
     *
     * @return $this
     */
    public function writeCell($value, ?array $styles = null): SheetTemplate
    {
        $this->sheetWriter->writeCell($value, $styles);

        return $this;
    }

}
