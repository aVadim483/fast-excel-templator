<?php

namespace avadim\FastExcelTemplator;

use avadim\FastExcelHelper\Helper;

class TableSection
{
    /** @var array  */
    public array $tplRange = [];

    /** @var array  */
    protected array $trgRange = [];

    protected Sheet $sheet;

    protected ?RowTemplateCollection $rowTemplates;

    /**
     * @param $sheet
     * @param string $range
     * @param bool|null $loadTemplates
     */
    public function __construct($sheet, string $range, ?bool $loadTemplates = false)
    {
        $this->sheet = $sheet;
        $this->tplRange = self::_rangeArray($range);
        if ($loadTemplates) {
            $this->rowTemplates = new RowTemplateCollection();
            for ($rowNum = $this->tplRange['min_row_num']; $rowNum <= $this->tplRange['max_row_num']; $rowNum++) {
                $row = $sheet->getRowTemplate($rowNum);
                $this->rowTemplates->addRowTemplate($row, $rowNum);
            }
        }
    }

    /**
     * @param string $range
     *
     * @return array
     */
    public static function _rangeArray(string $range): array
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

    public function _appendCol($colLetter, &$range)
    {
        $range['max_col_letter'] = $colLetter;
        $range['max_cell'] = $colLetter . $range['max_row_num'];
        $range['col_letters'][] = $colLetter;
    }

    /**
     * @param array $rowTemplates
     *
     * @return void
     */
    public function applyRowTemplates(array $rowTemplates)
    {
        $this->rowTemplates = new RowTemplateCollection();
        foreach ($rowTemplates as $rowNum => $rowTemplate) {
            if ($rowNum >= $this->tplRange['min_row_num'] && $rowNum <= $this->tplRange['max_row_num']) {
                $this->rowTemplates->addRowTemplate($rowTemplate, $rowNum);
            }
        }
    }

    /**
     * @param string|null $colSource
     * @param string|array $colTarget
     * @param bool|null $checkMerge
     *
     * @return $this
     */
    public function cloneColumn(?string $colSource, $colTarget, ?bool $checkMerge = false): TableSection
    {
        $colTarget = Helper::colLetterRange($colTarget);
        foreach ($colTarget as $col) {
            $this->_appendCol($col, $this->tplRange);
        }
        if ($colSource) {
            $this->rowTemplates->cloneCell($colSource, $colTarget, $checkMerge);
        }

        return $this;
    }

    /**
     * Add additional column to the right
     *
     * @return $this
     */
    public function addColumn(?int $num = 1): TableSection
    {
        for ($i = 1; $i <= $num; $i++) {
            $colTarget = Helper::colLetter($this->tplRange['max_col_num'] + $i);
            $this->cloneColumn(null, $colTarget);
        }

        return $this;
    }

    /**
     * Clone last column and add one to the right
     *
     * @return $this
     */
    public function appendColumn(?int $num = 1): TableSection
    {
        $colSource = Helper::colLetter($this->tplRange['max_col_num']);
        for ($i = 1; $i <= $num; $i++) {
            $colTarget = Helper::colLetter($this->tplRange['max_col_num'] + $i);
            $this->cloneColumn($colSource, $colTarget, true);
        }

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
        $row = $this->rowTemplates->next();
        $cnt1 = count($this->tplRange['col_letters']);
        $cnt2 = count($rowData);
        if ($cnt1 > $cnt2) {
            $idx = array_slice($this->tplRange['col_letters'], 0, $cnt2);
            $rowData = array_combine($idx, array_values($rowData));
        }
        elseif ($cnt1 < $cnt2) {
            $val = array_slice($rowData, 0, $cnt1);
            $rowData = array_combine($this->tplRange['col_letters'], $val);
        }
        else {
            $rowData = array_combine($this->tplRange['col_letters'], array_values($rowData));
        }
        if (!$this->trgRange) {
            $this->sheet->replaceRow($this->tplRange['min_row_num'], $row, $rowData);
            $this->trgRange = $this->tplRange;
            $this->trgRange['max_row_num'] = $this->trgRange['min_row_num'];
        }
        elseif ($this->trgRange['max_row_num'] < $this->tplRange['max_row_num']) {
            $this->sheet->replaceRow(++$this->trgRange['max_row_num'], $row, $rowData);
        }
        else {
            $this->sheet->insertRow(++$this->trgRange['max_row_num'], $row, $rowData);
        }
        $this->trgRange['max_cell'] = $this->trgRange['max_col_letter'] . $this->trgRange['max_row_num'];
    }


    public function transferRows()
    {
        if ($this->sheet->lastReadRowNum < $this->tplRange['min_row_num'] - 1) {
            $this->sheet->transferRows($this->tplRange['min_row_num'] - 1);
        }
        while ($this->sheet->lastReadRowNum < $this->tplRange['max_row_num']) {
            $row = $this->rowTemplates->next();
            $this->sheet->lastReadRowNum = $row->attribute('r');
            $this->sheet->replaceRow($this->sheet->lastReadRowNum + $this->sheet->countInsertedRows, $row);
            if ($this->sheet->lastReadRowNum >= $this->tplRange['max_row_num']) {
                break;
            }
        }
        $this->sheet->transferRows($this->tplRange['max_row_num'], true);
    }

    /**
     * @return array
     */
    public function getColumns(): array
    {
        return $this->tplRange['col_letters'];
    }

    /**
     * @param string $columns
     *
     * @return string
     */
    public function colToRange(string $columns): string
    {
        if (preg_match('/^[a-z]+$/i', $columns)) {
            $colLetterMin = $colLetterMax = strtoupper($columns);
        }
        else {
            $arr = Helper::rangeArray($columns);
            $colLetterMin = $arr['min_col_letter'];
            $colLetterMax = $arr['max_col_letter'];
        }

        return $colLetterMin . $this->trgRange['min_row_num'] . ':' . $colLetterMax . $this->trgRange['max_row_num'];
    }

}