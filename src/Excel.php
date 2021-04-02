<?php

declare(strict_types=1);

namespace Ep\Excel;

use PhpOffice\PhpSpreadsheet\IOFactory;

final class Excel
{
    public function simpleRead(string $filePath, array $options = []): array
    {
        $options = $this->normalize($options);
        $sheet = IOFactory::load($filePath)->setActiveSheetIndex($options['sheet']);
        $data = [];
        for ($row = $options['startRow']; $row <= $sheet->getHighestRow(); $row++) {
            $item = [];
            $k = 0;
            for ($col = 1; $col <= $this->colToInt($sheet->getHighestColumn()); $col++, $k++) {
                $column = $options['columns'][$k] ?? $k;
                $item[$column] = $sheet->getCellByColumnAndRow($col, $row)->getFormattedValue();
            }
            $data[] = $item;
        }
        return $data;
    }

    private function normalize(array $options): array
    {
        $options['sheet'] ??= 0;
        $options['startRow'] ??= 2;
        $options['columns'] ??= [];

        return $options;
    }

    private function colToInt($col): int
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
