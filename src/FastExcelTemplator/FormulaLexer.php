<?php

namespace avadim\FastExcelTemplator;

use avadim\FastExcelHelper\Helper;

/**
 * Finds cell references inside the text of an Excel formula, leaving everything else untouched
 *
 * A formula captured from a template is stored in RC (relative) notation, so that a row template
 * can be re-inserted at any row and have its references re-based (see SheetTemplate::_cellFormula()).
 * Rewriting "=SUM(A1:A10)" means telling apart what in that text really is a reference from what
 * merely looks like one: a string literal, an error constant, a sheet name, a function name or a
 * number written in exponent form.
 *
 * The whole scan is a single PCRE pass. Every alternative that is NOT a reference is matched and
 * then thrown away with (*SKIP)(*FAIL): PCRE resumes after that text and never lets it match the
 * reference alternative that follows. Only the last alternative captures, so the callback runs on
 * candidates only. Keeping the scan inside PCRE instead of a PHP character loop is what makes this
 * about three times faster than the PHP tokenizer it replaced.
 *
 * Known limits (the previous implementation shared them):
 *  - whole-column and whole-row ranges ("A:A", "1:1") are left as is;
 *  - a spilled-range reference ("A1#") is left as is, so its '#' cannot be mistaken for an error.
 *
 * @internal This class is an implementation detail of the templator and may move
 *           to avadim/fast-excel-helper, where the address arithmetic already lives
 */
class FormulaLexer
{
    /** Sheet or workbook name may contain any non-ASCII letters, so match them as bytes */
    private const NAME_CHARS = 'A-Za-z0-9_.\x80-\xFF';

    private static ?string $pattern = null;

    /** Sheet/workbook prefixes and table references are opaque in both formula notations. */
    private const OPAQUE = "(?:\\[\\d+\\])?(?:'(?:[^']++|'')+'|["
        . self::NAME_CHARS . "]+(?::[" . self::NAME_CHARS . "]+)?)!"
        . '|[' . self::NAME_CHARS . ']*(?&brackets)';

    // An apostrophe escapes a special character in a structured reference.
    private const BRACKETS = "(?(DEFINE)(?<brackets>\\[(?:[^\\[\\]']++|'[\\[\\]#@']|(?&brackets))*\\]))";

    /**
     * The single pattern used for the scan, built once per process
     *
     * @return string
     */
    private static function pattern(): string
    {
        if (self::$pattern === null) {
            $skip = '(*SKIP)(*FAIL)';
            $name = '[' . self::NAME_CHARS . ']';

            self::$pattern = '~'
                // a string literal, doubled quotes included: "he said ""yes"""
                . '"(?:[^"]|"")*"' . $skip
                // an error constant; without this '#' would start a reference-looking mess
                . '|\#(?:REF|N/A|DIV/0|VALUE|NAME|NULL|NUM|SPILL|CALC|FIELD|BLOCKED|CONNECT|UNKNOWN|GETTING_DATA)[!?]?' . $skip
                . '|(?:' . self::OPAQUE . ')' . $skip
                // a number; matching it here is what keeps "1E3" from being read as cell E3
                . '|\d+(?:\.\d+)?(?:[Ee][-+]?\d+)?' . $skip
                // a function name, i.e. a name followed by '(' — SUM(, LOG10(, _xlfn.IFS(
                . '|' . $name . '+\s*\(' . $skip
                // and finally the reference itself: a single cell or a whole range as one token
                . '|(?<![' . self::NAME_CHARS . '$])'
                . '(?P<ref>\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?)'
                . '(?![' . self::NAME_CHARS . '(\#])'
                . self::BRACKETS . '~';
        }

        return self::$pattern;
    }

    /** Protect prefixes and structured references from Writer's R1C1 and function normalization. */
    public static function maskOpaqueReferences(string $formula, array &$masked): string
    {
        $prefix = '_fxt_';
        while (strpos($formula, $prefix) !== false) {
            $prefix .= '_';
        }
        // Skip string literals and genuine relative R1C1 tokens. Writer handles both itself.
        // Do not skip a sheet prefix such as R2C3! or a table name followed by brackets.
        $rc = 'R(?:\[-?\d+\]|\d*)C(?:\[-?\d+\]|\d*)';
        $pattern = '~"(?:[^"]|"")*"(*SKIP)(*FAIL)'
            . '|(?<![' . self::NAME_CHARS . '])' . $rc . '(?![' . self::NAME_CHARS . '\[!])(*SKIP)(*FAIL)'
            . '|(?:' . self::OPAQUE . ')' . self::BRACKETS . '~';
        return preg_replace_callback($pattern, static function (array $match) use (&$masked, $prefix): string {
            $key = $prefix . count($masked) . '_';
            // Keep the separator required by Writer's R1C1 regex after a sheet prefix.
            if (substr($match[0], -1) === '!') {
                $key .= '!';
            }
            $masked[$key] = $match[0];
            return $key;
        }, $formula) ?? $formula;
    }

    /**
     * Apply a callback to every cell reference of the formula
     *
     * The callback receives one reference ("B2", "$A$1", "A1:A10") and returns its replacement;
     * returning the argument unchanged leaves the formula intact.
     *
     * @param string $formula
     * @param callable $callback function(string $reference): string
     *
     * @return string
     */
    public static function replaceReferences(string $formula, callable $callback): string
    {
        if ($formula === '') {
            return $formula;
        }

        return preg_replace_callback(self::pattern(), static function (array $m) use ($callback) {
            // groups other than 'ref' never capture: they are dropped by (*SKIP)(*FAIL)
            if (!isset($m['ref']) || $m['ref'] === '' || !self::isInsideSheet($m['ref'])) {
                return $m[0];
            }

            return (string)$callback($m['ref']);
        }, $formula) ?? $formula;
    }

    /**
     * Check that a reference addresses a cell (or a range) that can exist on a sheet
     *
     * The pattern accepts anything shaped like an address, so "ZZZ9999999" reaches this point too.
     * Excel does not allow a defined name that looks like a valid address, hence the reverse also
     * holds: what cannot be an address is a name and must be left alone.
     *
     * @param string $reference
     *
     * @return bool
     */
    private static function isInsideSheet(string $reference): bool
    {
        foreach (explode(':', $reference) as $part) {
            if (!preg_match('/^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/', $part, $m)) {
                return false;
            }
            $rowNum = (int)$m[2];
            if ($rowNum < 1 || $rowNum > Helper::EXCEL_2007_MAX_ROW
                || Helper::colNumber($m[1]) > Helper::EXCEL_2007_MAX_COL) {
                return false;
            }
        }

        return true;
    }

    /**
     * Convert every reference of the formula from A1 to RC notation, relative to a base address
     *
     * A reference that cannot be converted (an address outside the sheet, an unparseable range)
     * is kept as it was: a formula that stays in A1 notation is still readable by Excel, while a
     * half-converted one would not be.
     *
     * @param string $formula Formula with or without the leading '='
     * @param string $baseAddress Address the relative offsets are counted from, e.g. 'C7'
     *
     * @return string Formula in RC notation, always with the leading '='
     */
    public static function convertA1toRC(string $formula, string $baseAddress): string
    {
        $result = self::replaceReferences($formula, static function (string $ref) use ($baseAddress) {
            $converted = Helper::A1toRC($ref, $baseAddress);

            return ($converted !== '') ? $converted : $ref;
        });

        return (isset($result[0]) && $result[0] === '=') ? $result : '=' . $result;
    }
}
