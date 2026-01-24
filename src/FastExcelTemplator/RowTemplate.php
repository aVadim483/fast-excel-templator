<?php

namespace avadim\FastExcelTemplator;

use avadim\FastExcelHelper\Helper;

class RowTemplate implements \Iterator
{
    protected ?SheetTemplate $sheetTemplate = null;

    protected array $cells = [];

    protected array $attributes = [];

    protected array $lastAddedCells = [];


    /**
     * RowTemplate constructor
     *
     * @param array|null $cellData
     */
    public function __construct(?array $cellData = [])
    {
        if ($cellData) {
            $this->setValues($cellData);
        }
    }

    /**
     * Return the current element
     *
     * @return mixed
     */
    #[\ReturnTypeWillChange]
    public function current()
    {
        return current($this->cells);
    }

    /**
     * Return the key of the current element
     *
     * @return mixed
     */
    #[\ReturnTypeWillChange]
    public function key()
    {
        return key($this->cells);
    }

    /**
     * Move forward to next element
     *
     * @return void
     */
    #[\ReturnTypeWillChange]
    public function next()
    {
        next($this->cells);
    }

    /**
     * Rewind the Iterator to the first element
     *
     * @return void
     */
    #[\ReturnTypeWillChange]
    public function rewind()
    {
        reset($this->cells);
    }

    /**
     * Checks if current position is valid
     *
     * @return bool
     */
    public function valid(): bool
    {
        return (bool)current($this->cells);
    }

    /**
     * Set sheet template
     *
     * @param SheetTemplate $sheetTemplate
     *
     * @return $this
     */
    public function setSheetTemplate(SheetTemplate $sheetTemplate): RowTemplate
    {
        $this->sheetTemplate = $sheetTemplate;

        return $this;
    }

    /**
     * Get sheet template
     *
     * @return SheetTemplate|null
     */
    public function getSheetTemplate(): ?SheetTemplate
    {
        return $this->sheetTemplate;
    }

    /**
     * Create cell internal structure
     *
     * @param string|int $colLetter
     * @param mixed $value
     * @param mixed|null $style
     *
     * @return array
     */
    protected function _createCell($colLetter, $value, $style = null): array
    {
        if ($colLetter) {
            if (is_numeric($colLetter)) {
                $colLetter = Helper::colLetter($colLetter);
            }
            $address = strtoupper($colLetter) . $this->rowNumber();
        }
        else {
            $address = null;
        }
        $cell = [
            'v' => $value,
            's' => 0,
            'f' => null,
            't' => 'GENERAL',
            'b' => null,
            '__sourceAddress' => $address,
        ];
        $sheet = $this->getSheetTemplate();
        if ($sheet && $style) {
            /** @var StyleManager $styleManager */
            $styleManager = $sheet->excel->styleManager;
            $style = StyleManager::normalize($style);
            $cell['s'] = $styleManager->addStyle($style);
        }

        return $cell;
    }

    /**
     * Add a cell to the row
     *
     * @param string $colLetter
     * @param mixed $value
     * @param mixed|null $style
     *
     * @return $this
     *
     * @example
     * $row->addCell('G', $value);
     */
    public function addCell(string $colLetter, $value, $style = null): RowTemplate
    {
        if (is_scalar($value)) {
            $cell = $this->_createCell($colLetter, $value, $style);
        }
        else {
            $cell = $value;
        }

        $this->cells[strtoupper($colLetter)] = $cell;

        return $this;
    }

    /**
     * Append a source cell to the end of the row
     *
     * @param string|null $colSource
     * @param int|null $number
     *
     * @return $this
     *
     * @example
     * $row->appendCell('E'); // Append cell like E to the end of the row
     * $row->appendCell('E', 2); // Append two cells like E to the end of the row
     */
    public function appendCell(?string $colSource = null, ?int $number = null): RowTemplate
    {
        if (!$colSource) {
            $colSource = array_key_last($this->cells);
        }
        if (!$number) {
            $colTarget = Helper::colLetter(Helper::colNumber($colSource) + 1);
        }
        else {
            $colTarget = [];
            for ($i = 0; $i < $number; $i++) {
                $colTarget[] = Helper::colLetter(Helper::colNumber($colSource) + 1 + $i);
            }
        }
        $this->cloneCell($colSource, $colTarget);

        return $this;
    }

    /**
     * Clone a cell
     *
     * @param string $colSource
     * @param string|string[] $colTarget
     * @param bool|null $checkMerge
     *
     * @return RowTemplate
     */
    public function cloneCell(string $colSource, $colTarget, ?bool $checkMerge = false): RowTemplate
    {
        if (preg_match('/^([a-z]+)(\d+)/i', $colSource, $m)) {
            $colSource = $m[1];
        }
        $colSource = strtoupper($colSource);
        if (isset($this->cells[$colSource])) {
            $colTarget = Helper::colLetterRange($colTarget);
            foreach ($colTarget as $col) {
                if (is_object($this->cells[$colSource])) {
                    $cell = clone $this->cells[$colSource];
                }
                else {
                    $cell = $this->cells[$colSource];
                    if (!$checkMerge && !empty($cell['__merged'])) {
                        unset($cell['__merged']);
                    }
                }
                $this->addCell($col, $cell);
            }
            $this->lastAddedCells = $colTarget;
        }

        return $this;
    }

    /**
     * Mark a cell as removed
     *
     * @param string $col
     *
     * @return $this
     *
     * @example
     * $row->removeCell('G');
     */
    public function removeCell(string $col): RowTemplate
    {
        $col = strtoupper($col);
        if (isset($this->cells[$col])) {
            $this->cells[$col]['__removed'] = 1;
        }

        return $this;
    }

    /**
     * Mark multiple cells as removed
     *
     * @param string[] $cols
     *
     * @return $this
     *
     * @example
     * $row->removeCells(['G', 'H']);
     */
    public function removeCells(array $cols): RowTemplate
    {
        foreach ($cols as $col) {
            $col = strtoupper($col);
            if (isset($this->cells[$col])) {
                $this->cells[$col]['__removed'] = 1;
            }
        }

        return $this;
    }

    /**
     * Get all cells of the row
     *
     * @return array
     */
    public function cells(): array
    {
        return $this->cells;
    }

    /**
     * Set row attributes
     *
     * @param array $attributes
     *
     * @return RowTemplate
     */
    public function setAttributes(array $attributes): RowTemplate
    {
        $this->attributes = $attributes;

        return $this;
    }

    /**
     * All source row attributes
     *
     * @return array
     */
    public function getAttributes(): array
    {
        return $this->attributes;
    }

    /**
     * Source row attribute
     *
     * @param string $name
     *
     * @return string|null
     */
    public function attribute($name): ?string
    {
        return $this->attributes[$name] ?? null;
    }

    /**
     * Source row number
     *
     * @return int|null
     */
    public function rowNumber(): ?int
    {
        $rowNum = $this->attribute('r');

        return $rowNum ? (int)$rowNum : null;
    }

    /**
     * Source row height
     *
     * @return string|null
     */
    public function rowHeight(): ?string
    {
        return $this->attribute('ht');
    }

    /**
     * Returns value of cell
     *
     * @param string $colLetter
     *
     * @return mixed|null
     */
    public function getValue(string $colLetter)
    {
        $colLetter = strtoupper($colLetter);

        return$this->cells[$colLetter]['v'] ?? null;
    }

    /**
     * Get values of all cells of the row
     *
     * @return array
     */
    public function getValues(): array
    {
        $values = [];
        foreach ($this->cells as $col => $cell) {
            $values[$col] = $this->getValue($col);
        }

        return $values;
    }

    /**
     * Set cell value
     *
     * @param string $colLetter
     * @param mixed $value
     *
     * @return $this
     *
     * @example
     * $row->setValue('G', '=SUM(A1:A10)');
     */
    public function setValue(string $colLetter, $value): RowTemplate
    {
        $colLetter = strtoupper($colLetter);
        if ($value && is_string($value) && $value[0] === '=') {
            $this->cells[$colLetter]['f'] = $value;
            $this->cells[$colLetter]['v'] = '';
        }
        else {
            if (!is_scalar($value)) {
                throw new Exception('Value for column "' . $colLetter . '" is not scalar');
            }
            $this->cells[$colLetter]['v'] = $value;
        }
        if (!isset($this->cells[$colLetter]['t'])) {
            $this->cells[$colLetter]['t'] = '';
        }

        return $this;
    }

    /**
     * Set multiple cell values
     *
     * @param array $values
     *
     * @return $this
     *
     * @example
     * $row->setValues(['G' => '=SUM(A1:A10)', 'H' => 'value']);
     */
    public function setValues(array $values): RowTemplate
    {
        foreach ($values as $colLetter => $value) {
            $this->setValue($colLetter, $value);
        }

        return $this;
    }

    /**
     * Assign a value to the last added cell
     *
     * @param mixed $value
     *
     * @return $this
     *
     * @example
     * $row->appendCell('E')->withValue('value');
     */
    public function withValue($value): RowTemplate
    {
        if ($this->lastAddedCells) {
            $colLetter = end($this->lastAddedCells);
            $this->setValue($colLetter, $value);
        }
        else {
            throw new Exception('There is no added cell to assign a value to');
        }

        return $this;
    }

    /**
     * Assign values to the last added cells
     *
     * @param array $values
     *
     * @return $this
     *
     * @example
     * $row->appendCell('E', 2)->withValues(['value1', 'value2']);
     */
    public function withValues(array $values): RowTemplate
    {
        if (!$this->lastAddedCells) {
            throw new Exception('There are no added cells to assign values to');
        }
        $num = -1;
        foreach ($this->lastAddedCells as $colLetter) {
            if (isset($values[++$num])) {
                $this->setValue($colLetter, $values[$num]);
            }
        }

        return $this;
    }

    /**
     * Set cell style
     *
     * @param string $colLetter
     * @param mixed $style
     *
     * @return $this
     */
    public function setCellStyle(string $colLetter, $style): RowTemplate
    {
        $colLetter = strtoupper($colLetter);
        $cell = $this->_createCell($colLetter, '', $style);
        if (!isset($this->cells[$colLetter])) {
            $this->addCell($colLetter, $cell, $style);
        }
        else {
            $this->cells[$colLetter]['s'] = $cell['s'];
        }

        return $this;
    }

    /**
     * Set row style
     *
     * @param mixed $style
     *
     * @return $this
     */
    public function setRowStyle($style): RowTemplate
    {
        $sheet = $this->getSheetTemplate();
        if ($sheet && $style) {
            /** @var StyleManager $styleManager */
            $styleManager = $sheet->excel->styleManager;
            $style = StyleManager::normalize($style);
            $styleIdx = $styleManager->addStyle($style);
            $attributes = $this->getAttributes();
            $attributes['s'] = $styleIdx;
            $attributes['customFormat'] = '1';
            $this->setAttributes($attributes);
        }

        return $this;
    }
}