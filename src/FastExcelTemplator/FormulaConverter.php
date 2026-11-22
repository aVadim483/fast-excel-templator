<?php

namespace avadim\FastExcelTemplator;

/** @internal Preserve Excel reference syntax around Writer's formula normalization. */
final class FormulaConverter extends \avadim\FastExcelWriter\FormulaConverter
{
    public function normalize(string $formula, $baseAddress): string
    {
        $masked = [];
        $formula = FormulaLexer::maskOpaqueReferences($formula, $masked);
        return strtr(parent::normalize($formula, $baseAddress), $masked);
    }
}
