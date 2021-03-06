<?php

declare(strict_types=1);

use Ep\Excel\Excel;

define('BASE_PATH', dirname(__DIR__, 3));

require(BASE_PATH . '/vendor/autoload.php');

$excel = new Excel;

$result = $excel->simpleRead(BASE_PATH . '/tests/Support/list-2.xlsx', ['columns' => ['id', 'name']]);

echo '<xmp>';
print_r($result);
