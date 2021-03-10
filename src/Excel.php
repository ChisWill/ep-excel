<?php

declare(strict_types=1);

namespace Ep\Excel;

use PhpOffice\PhpSpreadsheet\IOFactory;

final class Excel
{
    public function simpleRead(string $filePath): array
    {
        $spreadsheet = IOFactory::load($filePath);
        $sheet = $spreadsheet->getActiveSheet();

        $maxRow = $sheet->getHighestRow();
        $maxCol = $sheet->getHighestColumn();

        print_r($maxRow);
        echo ' - ';
        print_r($maxCol);
        die;
    }
}
