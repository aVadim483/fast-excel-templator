<?php

namespace avadim\FastExcelTemplator;

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

    public function current()
    {
        return current($this->domCells);
    }

    public function key()
    {
        return key($this->domCells);
    }

    public function next()
    {
        return next($this->domCells);
    }
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