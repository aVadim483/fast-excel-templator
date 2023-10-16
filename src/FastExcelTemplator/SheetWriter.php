<?php

namespace avadim\FastExcelTemplator;

use avadim\FastExcelHelper\Helper;
use avadim\FastExcelWriter\Interfaces\InterfaceSheetWriter;

class SheetWriter extends \avadim\FastExcelWriter\Sheet implements InterfaceSheetWriter
{
    protected array $fill = [];

    protected array $replace = [];


    public function setNode($rowIdx, $colIdx, $cell)
    {
        $this->cells['nodes'][$rowIdx][$colIdx] = $cell;
    }

    public function setFillValues(array $fillData)
    {
        $this->fill = $fillData;
    }


    public function setReplaceValues(array $replaceData)
    {
        $this->replace = [
            'keys' => array_keys($replaceData),
            'values' => array_values($replaceData),
        ];
    }

    /**
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
     * @param $cellAddress
     * @param $value
     * @param $styles
     * @param bool|null $merge
     * @param bool|null $changeCurrent
     *
     * @return array
     */
    protected function _setCellData($cellAddress, $value, $styles = null, ?bool $merge = false, ?bool $changeCurrent = false): array
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
        return parent::_setCellData($cellAddress, $value, $styles, $merge, $changeCurrent);
    }

    public function _writeToCellByIdx($cellAddress, $value, $styles = null)
    {
        $this->_setCellData($cellAddress, $value, $styles);
    }
}