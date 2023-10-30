<?php

namespace avadim\FastExcelTemplator;

class RowTemplateCollection implements \Iterator
{
    protected array $rowTemplates = [];
    protected int $pointer = 0;


    public function __construct(?array $rowData = [])
    {
        if ($rowData) {
            foreach ($rowData as $row) {
                $this->addRowTemplate($row);
            }
        }
    }

    public function addRowTemplate($row)
    {
        $this->rowTemplates[] = $row;
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