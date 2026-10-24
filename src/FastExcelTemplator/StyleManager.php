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
        if ($this->loaded) {
            return;
        }
        $this->elements = [];
        $reader = new Reader($this->excelTemplator->templateFile());
        $sectionName = null;
        $sections = ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'dxfs', 'tableStyles'];
        try {
            $reader->openZip('xl/styles.xml');
            while ($reader->read()) {
                if ($reader->depth === 1) {
                    $sectionName = $reader->nodeType === \XMLReader::ELEMENT && in_array($reader->name, $sections, true)
                        ? $reader->name : null;
                    continue;
                }
                if ($reader->nodeType !== \XMLReader::ELEMENT || $reader->depth !== 2 || $sectionName === null) {
                    continue;
                }
                $name = $reader->name;
                if (!in_array($name, ['numFmt', 'font', 'fill', 'border', 'xf', 'cellStyle', 'dxf', 'tableStyle'], true)) {
                    continue;
                }
                $tag = str_replace(' xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"', '', $reader->readOuterXml());
                $value = $name === 'dxf' ? $tag : ['tag' => $tag];
                if (in_array($name, ['xf', 'cellStyle', 'tableStyle', 'numFmt'], true)) {
                    foreach ($reader->getAllAttributes() as $attrName => $attrValue) {
                        $attrName = '_' . strtolower(preg_replace('/(?<!^)[A-Z]/', '_$0', $attrName));
                        $value[$attrName] = $attrValue;
                    }
                    foreach ($reader->expand()->childNodes as $child) {
                        if (!$child instanceof \DOMElement) {
                            continue;
                        }
                        foreach ($child->attributes as $key => $val) {
                            $attrName = $child->nodeName . '-' . strtolower(preg_replace('/(?<!^)[A-Z]/', '-$0', $key));
                            $value[$attrName] = $val;
                        }
                    }
                }
                // Only direct children of known sections are loaded: fonts inside dxf
                // must not become ordinary fonts, and each original index must survive.
                $this->loadElement($sectionName, $value);
            }
        } finally {
            $reader->close();
        }
        $this->loaded = true;
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
