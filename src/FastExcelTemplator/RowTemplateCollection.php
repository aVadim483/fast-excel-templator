<?php

namespace avadim\FastExcelTemplator;

use avadim\FastExcelHelper\Helper;

class RowTemplateCollection implements \Iterator
{
    protected ?SheetTemplate $sheet = null;
    /** @var RowTemplate[]  */
    protected array $rowTemplates = [];
    protected ?int $pointer = null;

    /** Number of templates already returned by next() since the last rewind(), see valid() */
    protected int $visited = 0;

    /** Independent insertion cycle; foreach and next() keep their existing contracts. */
    protected array $insertionKeys = [];
    protected int $insertionPosition = 0;


    /**
     * RowTemplateCollection constructor
     *
     * @param array|null $rowData
     * @param SheetTemplate|null $sheet Sheet the templates were captured from; some operations
     *                                  (cloneCell) need it, it can also be set later by setSheet()
     */
    public function __construct(?array $rowData = [], ?SheetTemplate $sheet = null)
    {
        if ($rowData) {
            foreach ($rowData as $num => $row) {
                $this->addRowTemplate($row, $num);
            }
        }
        if ($sheet) {
            $this->setSheet($sheet);
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
        $this->insertionKeys = [];
        $this->insertionPosition = 0;
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
            $this->insertionKeys = [];
            $this->insertionPosition = 0;
        }
    }

    /**
     * Select the next insertion template independently of foreach/next()/rewind(); insertRow() shares this cycle and copies the selected template, while add/delete restarts it.
     *
     * @return RowTemplate
     * @throws Exception If the collection is empty
     */
    public function nextTemplate(): RowTemplate
    {
        if (!$this->insertionKeys) {
            $this->insertionKeys = array_keys($this->rowTemplates);
        }
        if (!$this->insertionKeys) {
            throw new Exception('Cannot insert from an empty template collection');
        }
        $key = $this->insertionKeys[$this->insertionPosition];
        $this->insertionPosition = ($this->insertionPosition + 1) % count($this->insertionKeys);
        return $this->rowTemplates[$key];
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
        // copying the column width is a bonus on top of cloning the cells, so a collection
        // built without a sheet still clones instead of failing on an uninitialized property
        if ($this->sheet) {
            $sourceColIdx = Helper::colIndex($colSource);
            $colAttributes = $this->sheet->sheetWriter->getColAttributes();
            if (isset($colAttributes[$sourceColIdx]['width'])) {
                $this->sheet->sheetWriter->setColWidth($colTarget, $colAttributes[$sourceColIdx]['width']);
            }
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
     * Move forward with wrap-around for compatibility; insertRow() uses the independent nextTemplate() cycle.
     *
     * @return mixed
     */
    #[\ReturnTypeWillChange]
    public function next()
    {
        $this->visited++;
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
        $this->visited = 0;

        return reset($this->rowTemplates);
    }

    /**
     * Checks if current position is valid
     *
     * Since next() deliberately wraps around, the position alone can never end a foreach.
     * The counter reset by rewind() is what makes a foreach walk the collection exactly once.
     *
     * @return bool
     */
    public function valid(): bool
    {
        return $this->visited < count($this->rowTemplates) && !empty(current($this->rowTemplates));
    }

}
