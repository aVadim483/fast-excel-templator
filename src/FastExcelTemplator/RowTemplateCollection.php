<?php

namespace avadim\FastExcelTemplator;

use avadim\FastExcelHelper\Helper;

class RowTemplateCollection implements \Iterator
{
    protected Sheet $sheet;
    /** @var RowTemplate[]  */
    protected array $rowTemplates = [];
    protected ?int $pointer = null;


    /**
     * @param array|null $rowData
     */
    public function __construct(?array $rowData = [])
    {
        if ($rowData) {
            foreach ($rowData as $num => $row) {
                $this->addRowTemplate($row, $num);
            }
        }
    }

    public function setSheet($sheet)
    {
        $this->sheet = $sheet;
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
     * @param int $rowNum
     *
     * @return void
     */
    public function delRowTemplate(int $rowNum)
    {
        if (isset($this->rowTemplates[$rowNum])) {
            unset($this->rowTemplates[$rowNum]);
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
        $sourceColIdx = Helper::colIndex($colSource);
        $colAttributes = $this->sheet->sheetWriter->getColAttributes();
        if (isset($colAttributes[$sourceColIdx]['width'])) {
            $this->sheet->sheetWriter->setColWidth($colTarget, $colAttributes[$sourceColIdx]['width']);
        }

        return $this;
    }

    #[\ReturnTypeWillChange]
    public function current()
    {
        return current($this->rowTemplates);
    }

    #[\ReturnTypeWillChange]
    public function key()
    {
        return key($this->rowTemplates);
    }

    #[\ReturnTypeWillChange]
    public function next()
    {
        $result = next($this->rowTemplates);
        if ($result === false || $this->pointer === null) {
            $result = reset($this->rowTemplates);
            $this->pointer = key($this->rowTemplates);
        }

        return $result;
    }

    #[\ReturnTypeWillChange]
    public function rewind()
    {
        $this->pointer = array_key_first($this->rowTemplates);

        return reset($this->rowTemplates);
    }

    public function valid(): bool
    {
        $result = current($this->rowTemplates);

        return !empty($result);
    }

}