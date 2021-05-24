<?php
require '../vendor/autoload.php';

use App\Controllers\ExcelController;

$excel = new ExcelController();
$excel->validate("./sample-data/Type_B.xlsx", "Xlsx");