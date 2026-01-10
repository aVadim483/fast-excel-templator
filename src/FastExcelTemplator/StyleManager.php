<?php

namespace avadim\FastExcelTemplator;

use avadim\FastExcelTemplator\Excel as ExcelTemplator;

class StyleManager extends \avadim\FastExcelWriter\Style\StyleManager
{
    public ExcelTemplator $excelTemplator;

    public bool $loaded = false;

    public function loadStyles()
    {
        if (!$this->loaded) {
            $this->elements = [];
            $reader = new Reader($this->excelTemplator->templateFile());
            $reader->openZip('xl/styles.xml');
            $sectionName = null;
            while ($reader->read()) {
                if ($reader->nodeType === \XMLReader::ELEMENT) {
                    if (in_array($reader->name, ['fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles'])) {
                        $sectionName = $reader->name;
                        continue;
                    }
                    if (in_array($name = $reader->name, ['font', 'fill', 'border', 'xf', 'cellStyle'])) {
                        $tag = $reader->readOuterXml();
                        $tag = str_replace(' xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"', '', $tag);
                        $value = ['tag' => $tag];
                        if ($name === 'xf' || $name === 'cellStyle') {
                            foreach ($reader->getAllAttributes() as $attrName => $attrValue) {
                                // camelCase to snake_case
                                $attrName = '_' . strtolower(preg_replace('/(?<!^)[A-Z]/', '_$0', $attrName));
                                $value[$attrName] = $attrValue;
                            }
                        }
                        $this->addElement($sectionName, $value);
                    }
                }
            }
            $this->loaded = true;
        }
    }
}