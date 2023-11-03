<?php

namespace avadim\FastExcelTemplator;

use avadim\FastExcelHelper\Helper;

class TableTemplate
{
    protected Sheet $sheet;
    protected array $tplRange = [];
    protected array $tplRangeBody = [];
    protected array $tplRangeHeader = [];
    protected array $tplRangeFooter = [];
    protected array $trgRange = [];
    protected ?RowTemplateCollection $rowTemplates;
    protected ?TableTemplate $header = null;
    protected ?TableTemplate $footer = null;
    protected bool $headerWritten = false;

    /**
     * @param $sheet
     * @param string $range
     * @param string|null $header
     * @param string|null $footer
     */
    public function __construct($sheet, string $range, ?string $header = null, ?string $footer = null)
    {
        $this->sheet = $sheet;
        $this->rowTemplates = new RowTemplateCollection();
        $this->tplRangeBody = $this->tplRange = $this->_rangeArray($range);
        if ($header) {
            $this->headerRange($header);
        }
        if ($footer) {
            $this->footerRange($footer);
        }
        for ($rowNum = $this->tplRangeBody['min_row_num']; $rowNum <= $this->tplRangeBody['max_row_num']; $rowNum++) {
            $row = $sheet->getRowTemplate($rowNum);
            $this->rowTemplates->addRowTemplate($row);
        }
    }

    /**
     * @param string $range
     *
     * @return array
     */
    protected function _rangeArray(string $range): array
    {
        static $rangeArray = [];

        if (isset($rangeArray[$range])) {
            $result = $rangeArray[$range];
        }
        else {
            $result = Helper::rangeArray($range);
            $result['col_letters'] = Helper::colLetterRange($result['min_col_letter'] . '-' . $result['max_col_letter']);
            $rangeArray[$range] = $result;
        }

        return $result;
    }

    /**
     * @param string $header
     *
     * @return $this
     */
    public function headerRange(string $header): TableTemplate
    {
        $this->tplRangeHeader = $this->_rangeArray($header);
        if (!isset($this->tplRangeBody['min_row_num'])) {
            $this->tplRangeBody = $this->tplRangeHeader;
            $this->tplRangeBody['min_row_num'] = $this->tplRangeHeader['max_row_num'] + 1;
        }
        elseif ($this->tplRangeBody['min_row_num'] <= $this->tplRangeHeader['max_row_num']) {
            $this->tplRangeBody['min_row_num'] = $this->tplRangeHeader['max_row_num'] + 1;
        }
        $this->header = new TableTemplate($this->sheet, $header);

        return $this;
    }

    /**
     * @param string $footer
     *
     * @return $this
     */
    public function footerRange(string $footer): TableTemplate
    {
        $this->tplRangeFooter = $this->_rangeArray($footer);
        if (!isset($this->tplRangeBody['max_row_num'])) {
            $this->tplRangeBody = $this->tplRangeFooter;
            $this->tplRangeBody['max_row_num'] = $this->tplRangeFooter['min_row_num'] - 1;
        }
        elseif ($this->tplRangeBody['max_row_num'] >= $this->tplRangeFooter['min_row_num']) {
            $this->tplRangeBody['max_row_num'] = $this->tplRangeFooter['min_row_num'] - 1;
        }
        $this->footer = new TableTemplate($this->sheet, $footer);

        return $this;
    }

    public function _appendCol($colLetter, &$range)
    {
        $range['max_col_letter'] = $colLetter;
        $range['max_cell'] = $colLetter . $range['max_row_num'];
        $range['col_letters'][] = $colLetter;
    }

    /**
     * @param string|null $colSource
     * @param $colTarget
     *
     * @return $this
     */
    public function cloneColumn(?string $colSource, $colTarget): TableTemplate
    {
        $colTarget = Helper::colLetterRange($colTarget);
        foreach ($colTarget as $col) {
            $this->_appendCol($col, $this->tplRange);
            $this->_appendCol($col, $this->tplRangeBody);
            if ($this->tplRangeHeader) {
                $this->_appendCol($col, $this->tplRangeBody);
            }
            if ($this->tplRangeFooter) {
                $this->_appendCol($col, $this->tplRangeFooter);
            }
        }
        if ($colSource) {
            $this->rowTemplates->cloneCell($colSource, $colTarget);
            if ($this->header) {
                $this->header->cloneColumn($colSource, $colTarget);
            }
            if ($this->footer) {
                $this->footer->cloneColumn($colSource, $colTarget);
            }
        }

        return $this;
    }

    /**
     * @param string|null $colSource
     *
     * @return $this
     */
    public function appendColumn(?string $colSource = null): TableTemplate
    {
        $colTarget = Helper::colLetter($this->tplRangeBody['max_col_num'] + 1);
        $this->cloneColumn($colSource, $colTarget);

        return $this;
    }

    /**
     * @return RowTemplate|void
     */
    public function nextRowTemplate()
    {
        return $this->rowTemplates->next();
    }

    /**
     * @param $rowData
     *
     * @return void
     */
    public function writeRow($rowData)
    {
        if (!$this->headerWritten) {
            $this->header->transferRows();
            $this->headerWritten = true;
        }
        $row = $this->rowTemplates->next();
        $cnt1 = count($this->tplRangeBody['col_letters']);
        $cnt2 = count($rowData);
        if ($cnt1 > $cnt2) {
            $idx = array_slice($this->tplRangeBody['col_letters'], 0, $cnt2);
            $rowData = array_combine($idx, array_values($rowData));
        }
        elseif ($cnt1 < $cnt2) {
            $val = array_slice($rowData, 0, $cnt1);
            $rowData = array_combine($this->tplRangeBody['col_letters'], $val);
        }
        else {
            $rowData = array_combine($this->tplRangeBody['col_letters'], array_values($rowData));
        }
        if (!$this->trgRange) {
            $this->sheet->replaceRow($this->tplRangeBody['min_row_num'], $row, $rowData);
            $this->trgRange = $this->tplRangeBody;
            $this->trgRange['max_row_num'] = $this->trgRange['min_row_num'];
        }
        elseif ($this->trgRange['max_row_num'] < $this->tplRangeBody['max_row_num']) {
            $this->sheet->replaceRow(++$this->trgRange['max_row_num'], $row, $rowData);
        }
        else {
            $this->sheet->insertRow(++$this->trgRange['max_row_num'], $row, $rowData);
        }
        $this->trgRange['max_cell'] = $this->trgRange['max_col_num'] . $this->trgRange['max_row_num'];
    }

    public function writeRowArray($rowData)
    {

    }

    public function writeRowPattern($rowData)
    {

    }

    public function transferRows()
    {
        if ($this->sheet->lastReadRowNum < $this->tplRangeBody['min_row_num']) {
            $this->sheet->transferRows($this->tplRangeBody['min_row_num'] - 1);
        }
        while ($this->sheet->lastReadRowNum < $this->tplRangeBody['max_row_num']) {
            $row = $this->rowTemplates->next();
            $this->sheet->replaceRow($this->sheet->countInsertedRows + 1, $row);
        }
    }

}