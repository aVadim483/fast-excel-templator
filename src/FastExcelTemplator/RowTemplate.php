<?php

namespace avadim\FastExcelTemplator;

use avadim\FastExcelHelper\Helper;

class RowTemplate implements \Iterator
{
    protected array $domCells = [];

    protected array $attributes = [];

    public function __construct(?array $cellData = [])
    {
        if ($cellData) {
            $this->setValues($cellData);
        }
    }

    public function addCell($colLetter, $cell)
    {
        $this->domCells[$colLetter] = $cell;
    }

    /**
     * @param string|null $colSource
     *
     * @return void
     */
    public function appendCell(?string $colSource = null)
    {
        if (!$colSource) {
            $colSource = array_key_last($this->domCells);
        }
        $colTarget = Helper::colLetter(Helper::colNumber($colSource) + 1);
        $this->cloneCell($colSource, $colTarget);
    }

    /**
     * @param string $colSource
     * @param $colTarget
     * @param bool|null $checkMerge
     *
     * @return void
     */
    public function cloneCell(string $colSource, $colTarget, ?bool $checkMerge = false)
    {
        if (preg_match('/^([a-z]+)(\d+)/i', $colSource, $m)) {
            $colSource = $m[1];
        }
        $colSource = strtoupper($colSource);
        if (isset($this->domCells[$colSource])) {
            $colTarget = Helper::colLetterRange($colTarget);
            foreach ($colTarget as $col) {
                if (is_object($this->domCells[$colSource])) {
                    $cell = clone $this->domCells[$colSource];
                }
                else {
                    $cell = $this->domCells[$colSource];
                    if (!$checkMerge && !empty($cell['__merged'])) {
                        unset($cell['__merged']);
                    }
                }
                $this->addCell($col, $cell);
            }
        }
    }

    /**
     * @param array $attributes
     *
     * @return void
     */
    public function setAttributes(array $attributes)
    {
        $this->attributes = $attributes;
    }

    /**
     * @param $name
     *
     * @return string|null
     */
    public function attribute($name): ?string
    {
        return $this->attributes[$name] ?? null;
    }

    #[\ReturnTypeWillChange]
    public function current()
    {
        return current($this->domCells);
    }

    #[\ReturnTypeWillChange]
    public function key()
    {
        return key($this->domCells);
    }

    #[\ReturnTypeWillChange]
    public function next()
    {
        return next($this->domCells);
    }

    #[\ReturnTypeWillChange]
    public function rewind()
    {
        return reset($this->domCells);
    }

    public function valid(): bool
    {
        return (bool)current($this->domCells);
    }

    public function setValue($colLetter, $value): RowTemplate
    {
        if ($value && is_string($value) && $value[0] === '=') {
            $this->domCells[$colLetter]['f'] = $value;
            $this->domCells[$colLetter]['v'] = '';
        }
        else {
            $this->domCells[$colLetter]['v'] = $value;
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