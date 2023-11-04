<?php

namespace avadim\FastExcelTemplator;

class RowTemplateCollection implements \Iterator
{
    /** @var RowTemplate[]  */
    protected array $rowTemplates = [];
    protected int $pointer = 0;


    /**
     * @param array|null $rowData
     */
    public function __construct(?array $rowData = [])
    {
        if ($rowData) {
            foreach ($rowData as $row) {
                $this->addRowTemplate($row);
            }
        }
    }

    /**
     * @param $row
     * @param int|null $rowNum
     *
     * @return void
     */
    public function addRowTemplate($row, ?int $rowNum = 0)
    {
        if ($rowNum) {
            $this->rowTemplates[$rowNum] = $row;
        }
        else {
            $this->rowTemplates[] = $row;
        }
    }

    /**
     * @param string $colSource
     * @param $colTarget
     * @param bool|null $checkMerge
     *
     * @return $this
     */
    public function cloneCell(string $colSource, $colTarget, ?bool $checkMerge = false): RowTemplateCollection
    {
        foreach ($this->rowTemplates as $rowTemplate) {
            $rowTemplate->cloneCell($colSource, $colTarget, $checkMerge);
        }

        return $this;
    }

    public function current()
    {
        return $this->rowTemplates[$this->pointer];
    }

    public function key()
    {
        return $this->pointer;
    }

    public function next()
    {
        if ($this->pointer >= count($this->rowTemplates)) {
            $this->pointer = 0;
        }
        return $this->rowTemplates[$this->pointer++];
    }
    public function rewind()
    {
        $this->pointer = 0;

        return $this->rowTemplates[$this->pointer];
    }
    public function valid(): bool
    {
        return isset($this->rowTemplates[$this->pointer]);
    }

}