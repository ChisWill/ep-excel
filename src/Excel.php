<?php

declare(strict_types=1);

namespace Ep\Excel;

use PhpOffice\PhpSpreadsheet\IOFactory;

final class Excel
{
    public function simpleRead(string $filePath, array $options = []): array
    {
        $this->normalize($options);

        $sheet = IOFactory::load($filePath)->setActiveSheetIndex($options['sheet']);

        $maxRow = $sheet->getHighestRow();
        $maxCol = $this->colToInt($sheet->getHighestColumn());
        for ($row = $options['startRow']; $row <= $maxRow; $row++) {
            $item = [];
            for ($key = 0, $col = 1; $col <= $maxCol; $key++, $col++) {
                $item[$options['columns'][$key] ?? $key] = $sheet->getCellByColumnAndRow($col, $row)->getFormattedValue();
            }
            $data[] = $item;
        }
        return $data;
    }

    private function normalize(array &$options): array
    {
        $options['sheet'] ??= 0;
        $options['startRow'] ??= 2;
        $options['columns'] ??= [];

        return $options;
    }

    private function colToInt(string $col): int
    {
        $pieces = str_split($col);
        $power = count($pieces) - 1;
        $sum = 1;
        foreach ($pieces as $v) {
            $sum += (ord($v) - 64) * pow(26, $power);
            $power--;
        }
        return --$sum;
    }
}
