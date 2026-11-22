<?php

declare(strict_types=1);

namespace avadim\FastExcelTemplator;

/** @internal Source rows and their output positions, in output order. */
final class RowMapping
{
    /** @var array<int, array{int, int, int, int}> */
    private array $runs = [];

    public function record(int $source, int $last, int $target): void
    {
        $previous = array_key_last($this->runs);
        if ($previous !== null) {
            [$from, $through, $out, $outThrough] = $this->runs[$previous];
            if ($target === $outThrough + 1) {
                if ($from === $through && $through === $source && $source === $last) {
                    // Consecutive copies of one source row.
                    $this->runs[$previous][3] = $target;
                    return;
                }
                if ($through + 1 === $source && $through - $from === $outThrough - $out) {
                    $this->runs[$previous] = [$from, $last, $out, $target + $last - $source];
                    return;
                }
            }
        }
        $this->runs[] = [$source, $last, $target, $target + $last - $source];
    }

    /** @return array<int, array{int, int, ?int}> Output bounds and uniform row displacement, if any. */
    public function project(int $first, int $last): array
    {
        $result = [];
        foreach ($this->runs as [$from, $through, $out, $outThrough]) {
            if ($through < $first || $from > $last) {
                continue;
            }
            $start = $out + max(0, $first - $from);
            $end = $outThrough - max(0, $through - $last);
            $shift = $through - $from === $outThrough - $out ? $out - $from : null;
            $previous = array_key_last($result);
            if ($previous !== null && $start === $result[$previous][1] + 1) {
                $result[$previous][1] = $end;
                if ($result[$previous][2] !== $shift) {
                    $result[$previous][2] = null;
                }
            } else {
                $result[] = [$start, $end, $shift];
            }
        }
        return $result;
    }
}
