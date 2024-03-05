<?php

namespace avadim\FastExcelTemplator;

use avadim\FastExcelHelper\Helper;

class TableTemplate
{
    /** @var array  */
    public array $tplRange = [];

    protected Sheet $sheet;
    protected ?TableSection $header = null;
    protected ?TableSection $body = null;
    protected ?TableSection $footer = null;
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
        $this->body = new TableSection($this->sheet, $range);
        $this->tplRange = $this->body->tplRange;
        if ($header) {
            $this->header = new TableSection($this->sheet, $header);
            if ($this->body->tplRange['min_row_num'] <= $this->header->tplRange['max_row_num']) {
                $this->body->tplRange['min_row_num'] = $this->header->tplRange['max_row_num'] + 1;
            }
            if ($this->tplRange['min_row_num'] > $this->header->tplRange['min_row_num']) {
                $this->tplRange['min_row_num'] = $this->header->tplRange['min_row_num'];
            }
        }
        if ($footer) {
            $this->footer = new TableSection($this->sheet, $footer);
            if ($this->body->tplRange['max_row_num'] <= $this->footer->tplRange['min_row_num']) {
                $this->body->tplRange['max_row_num'] = $this->footer->tplRange['min_row_num'] - 1;
            }
            if ($this->tplRange['max_row_num'] < $this->header->tplRange['max_row_num']) {
                $this->tplRange['max_row_num'] = $this->header->tplRange['max_row_num'];
            }
        }
        $rowTemplates = $this->sheet->getRowTemplates($this->tplRange['min_row_num'], $this->tplRange['max_row_num']);
        if ($this->header) {
            $this->header->applyRowTemplates($rowTemplates);
        }
        $this->body->applyRowTemplates($rowTemplates);
        if ($this->footer) {
            $this->footer->applyRowTemplates($rowTemplates);
        }
    }

    /**
     * @return TableSection|null
     */
    public function header(): ?TableSection
    {
        return $this->header;
    }

    /**
     * @return TableSection|null
     */
    public function body(): ?TableSection
    {
        return $this->body;
    }

    /**
     * @return TableSection|null
     */
    public function footer(): ?TableSection
    {
        return $this->footer;
    }

    /**
     * @param string|null $colSource
     * @param string|array $colTarget
     * @param bool|null $checkMerge
     *
     * @return $this
     */
    public function cloneColumn(?string $colSource, $colTarget, ?bool $checkMerge = false): TableTemplate
    {
        $colTarget = Helper::colLetterRange($colTarget);
        if ($this->header) {
            $this->header->cloneColumn($colSource, $colTarget, $checkMerge);
        }
        if ($this->body) {
            $this->body->cloneColumn($colSource, $colTarget, $checkMerge);
        }
        if ($this->footer) {
            $this->footer->cloneColumn($colSource, $colTarget, $checkMerge);
        }

        return $this;
    }

    /**
     * Add additional column to the right
     *
     * @param int|null $num
     *
     * @return $this
     */
    public function addColumn(?int $num = 1): TableTemplate
    {
        if ($this->header) {
            $this->header->addColumn($num);
        }
        if ($this->body) {
            $this->body->addColumn($num);
        }
        if ($this->footer) {
            $this->footer->addColumn($num);
        }

        return $this;
    }

    /**
     * Clone last column and add one to the right
     *
     * @return $this
     */
    public function appendColumn(?int $num = 1): TableTemplate
    {
        if ($this->header) {
            $this->header->appendColumn($num);
        }
        if ($this->body) {
            $this->body->appendColumn($num);
        }
        if ($this->footer) {
            $this->footer->appendColumn($num);
        }

        return $this;
    }

    /**
     * @param $rowData
     *
     * @return void
     */
    public function writeRow($rowData)
    {
        if (!$this->headerWritten) {
            if ($this->header) {
                $this->header->transferRows();
            }
            $this->headerWritten = true;
        }
        $this->body->writeRow($rowData);
    }


    public function transferRows()
    {
        if ($this->header) {
            $this->header->transferRows();
        }
        $this->body->transferRows();
        if ($this->footer) {
            $this->footer->transferRows();
        }
    }

    /**
     * @return array
     */
    public function getColumns(): array
    {
        return $this->body->getColumns();
    }

    /**
     * @param string $columns
     *
     * @return string
     */
    public function colToRange(string $columns): string
    {
        return $this->body->colToRange($columns);
    }
}