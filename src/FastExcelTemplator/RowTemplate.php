<?php

namespace avadim\FastExcelTemplator;

use avadim\FastExcelHelper\Helper;

class RowTemplate implements \Iterator
{
    protected array $cells = [];

    protected array $attributes = [];

    protected array $lastAddedCells = [];


    public function __construct(?array $cellData = [])
    {
        if ($cellData) {
            $this->setValues($cellData);
        }
    }

    #[\ReturnTypeWillChange]
    public function current()
    {
        return current($this->cells);
    }

    #[\ReturnTypeWillChange]
    public function key()
    {
        return key($this->cells);
    }

    #[\ReturnTypeWillChange]
    public function next()
    {
        return next($this->cells);
    }

    #[\ReturnTypeWillChange]
    public function rewind()
    {
        return reset($this->cells);
    }

    public function valid(): bool
    {
        return (bool)current($this->cells);
    }

    /**
     * @param string $colLetter
     * @param $cell
     *
     * @return $this
     */
    public function addCell(string $colLetter, $cell): RowTemplate
    {
        $this->cells[strtoupper($colLetter)] = $cell;

        return $this;
    }

    /**
     * @param string|null $colSource
     * @param int|null $number
     *
     * @return $this
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

    public function removeCell(string $col): RowTemplate
    {
        $col = strtoupper($col);
        if (isset($this->cells[$col])) {
            $this->cells[$col]['__removed'] = 1;
        }

        return $this;
    }

    /**
     * @param string[] $cols
     *
     * @return $this
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
     * @return array
     */
    public function cells(): array
    {
        return $this->cells;
    }

    /**
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
     * @param $name
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
     * @param string $colLetter
     *
     * @return mixed|null
     */
    public function getValue(string $colLetter)
    {
        $colLetter = strtoupper($colLetter);

        return$this->cells[$colLetter]['v'] ?? null;
    }

    public function getValues(): array
    {
        $values = [];
        foreach ($this->cells as $col => $cell) {
            $values[$col] = $this->getValue($col);
        }

        return $values;
    }

    /**
     * @param string $colLetter
     * @param mixed $value
     *
     * @return $this
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
     * @param array $values
     *
     * @return $this
     */
    public function setValues(array $values): RowTemplate
    {
        foreach ($values as $colLetter => $value) {
            $this->setValue($colLetter, $value);
        }

        return $this;
    }

    /**
     * @param $value
     *
     * @return $this
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
     * @param array $values
     *
     * @return $this
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
}