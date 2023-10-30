<?php

namespace avadim\FastExcelTemplator;

use avadim\FastExcelHelper\Helper;

class TableTemplate
{
    protected Sheet $sheet;
    protected array $tplRange;
    protected array $trgRange = [];
    protected ?RowTemplateCollection $rowTemplates;

    /**
     * @param $sheet
     * @param $range
     *
     * @return void
     */
    public function __construct($sheet, $range)
    {
        $this->sheet = $sheet;
        $this->rowTemplates = new RowTemplateCollection();
        $this->tplRange = Helper::rangeArray($range);
        $this->tplRange['col_letters'] = Helper::colLetterRange($this->tplRange['min_col_letter'] . '-' . $this->tplRange['max_col_letter']);
        for ($rowNum = $this->tplRange['min_row_num']; $rowNum <= $this->tplRange['max_row_num']; $rowNum++) {
            $row = $sheet->getRowTemplate($rowNum);
            $this->rowTemplates->addRowTemplate($row);
        }
    }

    public function appendColumn(?string $asCol)
    {

    }
    public function writeRow($rowData)
    {
        $row = $this->rowTemplates->next();
        $rowData = array_combine($this->tplRange['col_letters'], array_values($rowData));
        if (!$this->trgRange) {
            $this->sheet->replaceRow($this->tplRange['min_row_num'], $row, $rowData);
            $this->trgRange = $this->tplRange;
            $this->trgRange['max_row_num'] = $this->trgRange['min_row_num'];
        }
        elseif ($this->trgRange['max_row_num'] < $this->tplRange['max_row_num']) {
            $this->sheet->replaceRow(++$this->trgRange['min_row_num'], $row, $rowData);
        }
        else {
            $this->sheet->insertRow(++$this->trgRange['min_row_num'], $row, $rowData);
        }
        $this->trgRange['max_cell'] = $this->trgRange['max_col_num'] . $this->trgRange['max_row_num'];
    }

    public function writeRowArray($rowData)
    {

    }

    public function writeRowPattern($rowData)
    {

    }

}