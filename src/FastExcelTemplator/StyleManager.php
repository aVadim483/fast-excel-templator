<?php

namespace avadim\FastExcelTemplator;

use avadim\FastExcelTemplator\Excel as ExcelTemplator;

class StyleManager extends \avadim\FastExcelWriter\Style\StyleManager
{
    public ExcelTemplator $excelTemplator;

    public bool $loaded = false;

    /**
     * Load styles from template
     */
    public function loadStyles()
    {
        if (!$this->loaded) {
            $this->elements = [];
            $reader = new Reader($this->excelTemplator->templateFile());
            $reader->openZip('xl/styles.xml');
            $sectionName = null;
            while ($reader->read()) {
                if ($reader->nodeType === \XMLReader::ELEMENT) {
                    if (in_array($reader->name, ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'dxfs', 'tableStyles'])) {
                        $sectionName = $reader->name;
                        continue;
                    }
                    if (in_array($name = $reader->name, ['numFmt', 'font', 'fill', 'border', 'xf', 'cellStyle'])) {
                        $tag = $reader->readOuterXml();
                        $tag = str_replace(' xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"', '', $tag);
                        $value = ['tag' => $tag];
                        if ($name === 'xf' || $name === 'cellStyle' || $name === 'tableStyle' || $name === 'numFmt') {
                            foreach ($reader->getAllAttributes() as $attrName => $attrValue) {
                                // camelCase to snake_case
                                $attrName = '_' . strtolower(preg_replace('/(?<!^)[A-Z]/', '_$0', $attrName));
                                $value[$attrName] = $attrValue;
                            }
                            $node = $reader->expand();
                            if ($node->hasChildNodes()) {
                                foreach ($node->childNodes as $child) {
                                    $name = $child->nodeName;
                                    foreach ($child->attributes as $key => $val) {
                                        $attrName = $name . '-' . strtolower(preg_replace('/(?<!^)[A-Z]/', '-$0', $key));
                                        $value[$attrName] = $val;
                                    }
                                }
                            }
                        }
                        $this->loadElement($sectionName, $value);
                    }
                }
            }
            $this->loaded = true;
        }
    }

    /**
     * @param string $sectionName
     * @param string|array $value
     *
     * @return int
     */
    protected function loadElement(string $sectionName, $value): int
    {
        if (is_string($value)) {
            $key = $value;
        }
        elseif (isset($value['tag'])) {
            $key = $value['tag'];
        }
        else {
            ksort($value);
            $key = json_encode($value);
        }

        $index = isset($this->elements[$sectionName]) ? count($this->elements[$sectionName]) : 0;
        if (isset($this->elements[$sectionName][$key])) {
            $key = $index . ':' . $key;
        }
        if ($sectionName === 'numFmts') {
            if (isset($value['_num_fmt_id'])) {
                $index = (int)$value['_num_fmt_id'];
            }
            else {
                $index += 164;
            }
        }
        $this->elements[$sectionName][$key] = [
            'index' => $index,
            'value' => $value,
        ];
        $this->elementIndexes[$sectionName][$index] = $key;

        return $index;
    }

}