<?php

declare(strict_types=1);

namespace Ep\Excel;

use PhpOffice\PhpSpreadsheet\IOFactory;

final class Excel
{
    public function simpleRead(string $filePath): array
    {
        $sheet = IOFactory::load($filePath)->getActiveSheet();
        $maxRow = $sheet->getHighestRow();
        $maxCol = $sheet->getHighestColumn();

        $loadTh = $loadTd = 0;
        $title = [];
        $return = [];
        for ($row = 1; $row <= $maxRow; $row++) {
            for ($col = 1; $col <= $this->colToInt($maxCol); $col++) {
                $v = $sheet->getCellByColumnAndRow($col, $row)->getFormattedValue();

                //     $val = trim($sheet->getCellByColumnAndRow($col + 1, $row)->getValue());
                //     eval($evalIfstr1);
                //     if ($ifStr) {
                //         $loadTh = 1;
                //         $title[$col] = $val;
                //     }
                //     if ($loadTd === 1) {
                //         if (isset($title[$col])) {
                //             eval($evalCellStr);
                //         }
                //     }
                // }
                // if ($loadTd === 1) {
                //     eval($evalRowStr);
                // }
                // if ($loadTh === 1) {
                //     $loadTh = 0;
                //     $loadTd = 1;
            }
        }
        if (isset($options['rowCallback'])) {
            return [];
        } else {
            return $return;
        }
    }

    private function colToInt($col)
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

/**
 * 调试专用，可以传入任意多的变量进行打印查看
 */
function tes()
{
    $isCli = PHP_SAPI === 'cli';
    if (!$isCli && !in_array('Content-type:text/html;charset=utf-8', headers_list())) {
        header('Content-type:text/html;charset=utf-8');
    }
    global $_debugFunc;
    $_debugFunc = $_debugFunc ?: 'print_r';
    foreach (func_get_args() as $msg) {
        if ($isCli) {
            $_debugFunc($msg);
            echo PHP_EOL;
        } else {
            if ($_debugFunc === 'var_dump') {
                $_debugFunc($msg);
            } else {
                echo '<xmp>';
                $_debugFunc($msg);
                echo '</xmp>';
            }
        }
    }
}

/**
 * @see tes()
 */
function test()
{
    global $_debugFunc;
    $_debugFunc = 'print_r';
    tes(...func_get_args());
    exit;
}

/**
 * @see tes()
 */
function dump()
{
    global $_debugFunc;
    $_debugFunc = 'var_dump';
    tes(...func_get_args());
    exit;
}
