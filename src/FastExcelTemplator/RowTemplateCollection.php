<?php

namespace avadim\FastExcelTemplator;

use avadim\FastExcelHelper\Helper;

class RowTemplateCollection implements \Iterator
{
    protected SheetTemplate $sheet;
    /** @var RowTemplate[]  */
    protected array $rowTemplates = [];
    protected ?int $pointer = null;


    /**
     * RowTemplateCollection constructor
     *
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

    /**
     * Set sheet template
     *
     * @param SheetTemplate $sheet
     */
    public function setSheet($sheet)
    {
        $this->sheet = $sheet;
    }

    /**
     * Add row template to collection
     *
     * @param RowTemplate $row
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
     * Delete row template from collection
     *
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
     * Clone cell in all row templates of the collection
     *
     * @param string $colSource
     * @param string|array $colTarget
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

    /**
     * Return the current element
     *
     * @return mixed
     */
    #[\ReturnTypeWillChange]
    public function current()
    {
        return current($this->rowTemplates);
    }

    /**
     * Return the key of the current element
     *
     * @return mixed
     */
    #[\ReturnTypeWillChange]
    public function key()
    {
        return key($this->rowTemplates);
    }

    /**
     * Move forward to next element
     *
     * @return mixed
     */
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

    /**
     * Rewind the Iterator to the first element
     *
     * @return mixed
     */
    #[\ReturnTypeWillChange]
    public function rewind()
    {
        $this->pointer = array_key_first($this->rowTemplates);

        return reset($this->rowTemplates);
    }

    /**
     * Checks if current position is valid
     *
     * @return bool
     */
    public function valid(): bool
    {
        $result = current($this->rowTemplates);

        return !empty($result);
    }

}