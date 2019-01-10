<?php
/**
 * Created by PhpStorm.
 * User: yq05
 * Date: 2018/5/28
 * Time: 上午11:19
 */
require_once('autoloadconfig.php');  //判断是否在线
//require '/Applications/XAMPP/xamppfiles/htdocs/composer/vendor/autoload.php';
//require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1', 'Hello World !');

$writer = new Xlsx($spreadsheet);
$writer->save('hello world.xlsx');