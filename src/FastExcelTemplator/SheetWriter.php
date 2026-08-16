<?php

namespace avadim\FastExcelTemplator;

use avadim\FastExcelHelper\Helper;
use avadim\FastExcelWriter\Interfaces\InterfaceSheetWriter;
use avadim\FastExcelWriter\Style\Style;

class SheetWriter extends \avadim\FastExcelWriter\Sheet implements InterfaceSheetWriter
{
    protected array $fill = [];

    protected array $replace = [];


    /**
     * Set sheet views attributes
     *
     * The parent keeps plain attribute arrays in $sheetViews and wraps each of them into '_attr'
     * itself, so the '_attr' level of the captured structure is unwrapped here, at write time,
     * rather than inside the getter. The '_items' level (pane and selection nodes) is dropped on
     * purpose: the writer regenerates those from the freeze settings, which Excel::_importSheets()
     * transfers separately via setFreezeRows()/setFreezeColumns().
     *
     * @param array $attributes
     */
    public function _setSheetViewsAttributes(array $attributes)
    {
        $this->sheetViews = [$attributes['_attr'] ?? $attributes];
    }

    /**
     * Set sheet format properties attributes
     *
     * @param array $attributes
     */
    public function _setSheetFormatPrAttributes(array $attributes)
    {
        foreach ($attributes as $key => $val) {
            if (strpos($key, ':') === false) {
                $this->sheetFormatPr[$key] = $val;
            }
        }
    }

    /**
     * Set node
     *
     * @param int $rowIdx
     * @param int $colIdx
     * @param mixed $cell
     *
     * @return void
     */
    public function setNode($rowIdx, $colIdx, $cell)
    {
        $this->cells['nodes'][$rowIdx][$colIdx] = $cell;
    }

    /**
     * Set fill values
     *
     * @param array $fillData
     *
     * @return void
     */
    public function setFillValues(array $fillData)
    {
        $this->fill = $fillData;
    }

    /**
     * Set replace values
     *
     * @param array $replaceData
     *
     * @return void
     */
    public function setReplaceValues(array $replaceData)
    {
        $this->replace = [
            'keys' => array_keys($replaceData),
            'values' => array_values($replaceData),
        ];
    }

    /**
     * Set style index for cell
     *
     * @param string $cellAddress
     * @param int $styleIdx
     * @param string|null $numberFormatType
     */
    public function _setStyleIdx(string $cellAddress, int $styleIdx, ?string $numberFormatType = null)
    {
        if (preg_match('/^([A-Z]+)(\d+)/i', $cellAddress, $m)) {
            $colNum = Helper::colNumber($m[1]);
            $rowNum = (int)$m[2];
            if ($rowNum <= $this->rowCountWritten) {
                throw new Exception('Row number must be greater then written rows');
            }
            $this->cells['styles'][$rowNum - 1][$colNum - 1]['_xf_id'] = $styleIdx;
            if ($numberFormatType) {
                $this->cells['styles'][$rowNum - 1][$colNum - 1]['number_format_type'] = $numberFormatType;
            }
        }
    }

    /**
     * Set cell data internal handler
     *
     * @param string|array $cellAddress
     * @param mixed $value
     * @param mixed|null $styles
     * @param int|null $mergeFlag
     * @param bool|null $changeCurrent
     *
     * @return array
     */
    protected function _setCellData($cellAddress, $value, $styles = null, ?int $mergeFlag = 0, ?bool $changeCurrent = false): array
    {
        if ($value && is_string($value)) {
            if ($this->fill) {
                foreach ($this->fill as $key => $val) {
                    if ($value === $key) {
                        $value = $val;
                        break;
                    }
                }
            }
            if ($this->replace) {
                $value = str_replace($this->replace['keys'], $this->replace['values'], $value);
            }
        }
        if (is_array($value) && isset($value['value'])) {
            $result = parent::_setCellData($cellAddress, $value['value'], $styles, $mergeFlag, $changeCurrent);
            $this->cells['values'][$cellAddress['row_idx']][$cellAddress['col_idx']] = $value;
        }
        else {
            $result = parent::_setCellData($cellAddress, $value, $styles, $mergeFlag, $changeCurrent);
        }

        return $result;
    }

    /**
     * Write to cell by index
     *
     * @param array $cellAddress
     * @param mixed $value
     * @param mixed|null $styles
     *
     * @return void
     */
    public function _writeToCellByIdx($cellAddress, $value, $styles = null)
    {

        if (isset($value['t']) && $value['t'] === 'e' && isset($this->cells['nodes'][$cellAddress['row_idx']][$cellAddress['col_idx']])) {
            $cell = $this->cells['nodes'][$cellAddress['row_idx']][$cellAddress['col_idx']];
            $attributes = [];
            foreach ($cell->attributes as $attr) {
                if (!in_array($attr->nodeName, ['r', 's', 't'])) {
                    $attributes[$attr->nodeName] = $attr->nodeValue;
                }
            }
            $value = ['value' => $value['v'], 'attr' => $attributes];
        }
        $this->_setCellData($cellAddress, $value, $styles, false, true);
    }

    /**
     * Update merged cells range
     *
     * @param string $oldRange
     * @param string $newRange
     *
     * @return void
     */
    public function updateMergedCells($oldRange, $newRange)
    {
        if (isset($this->mergeCells[$oldRange])) {
            unset($this->mergeCells[$oldRange]);
        }
        $this->mergeCells($newRange, 1);
    }

    /**
     * Current row number
     *
     * @return int
     */
    public function currentRowNum(): int
    {
        return $this->currentRowIdx + 1;
    }

    /**
     * Set row attributes
     *
     * @param int $rowNum
     * @param array $attributes
     *
     * @return $this
     */
    public function setRowAttributes(int $rowNum, array $attributes): SheetWriter
    {
        if (isset($attributes['r'])) {
            unset($attributes['r']);
        }
        $this->rowAttributes[$rowNum - 1] = $attributes;

        return $this;
    }
}