<?php

namespace avadim\FastExcelTemplator;

use avadim\FastExcelHelper\Helper;

class RowTemplate implements \Iterator
{
    protected array $cells = [];

    protected array $attributes = [];

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
     * @param $colLetter
     * @param $cell
     *
     * @return $this
     */
    public function addCell($colLetter, $cell): RowTemplate
    {
        $this->cells[$colLetter] = $cell;

        return $this;
    }

    /**
     * @param string|null $colSource
     *
     * @return $this
     */
    public function appendCell(?string $colSource = null): RowTemplate
    {
        if (!$colSource) {
            $colSource = array_key_last($this->cells);
        }
        $colTarget = Helper::colLetter(Helper::colNumber($colSource) + 1);
        $this->cloneCell($colSource, $colTarget);

        return $this;
    }

    /**
     * @param string $colSource
     * @param $colTarget
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

    public function setValue($colLetter, $value): RowTemplate
    {
        if ($value && is_string($value) && $value[0] === '=') {
            $this->cells[$colLetter]['f'] = $value;
            $this->cells[$colLetter]['v'] = '';
        }
        else {
            $this->cells[$colLetter]['v'] = $value;
        }
        if (!isset($this->cells[$colLetter]['t'])) {
            $this->cells[$colLetter]['t'] = '';
        }

        return $this;
    }

    public function setValues(array $values): RowTemplate
    {
        foreach ($values as $colLetter => $value) {
            $this->setValue($colLetter, $value);
        }

        return $this;
    }
}