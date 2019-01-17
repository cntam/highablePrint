<?php
/**
 * Created by PhpStorm.
 * User: yq05
 * Date: 2018/5/28
 * Time: 上午11:19
 */

require_once ('autoloadconfig.php');  //判断是否在线
//require '/Applications/XAMPP/xamppfiles/htdocs/composer/vendor/autoload.php';
//require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

//use PhpOffice\PhpSpreadsheet\Spreadsheet;
//use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;



$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1', 'Hello World !');
$sheet->setCellValue('B2', '123456');

$sheet->getStyle('B2')->getNumberFormat()->applyFromArray( [ 'formatCode' => NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE ] );

$writer = new Xlsx($spreadsheet);
$writer->save('hello world.xlsx');


//    // Redirect output to a client’s web browser (Xlsx)
//    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
//    header('Content-Disposition: attachment;filename='."12");
//    header('Cache-Control: max-age=0');
//// If you're serving to IE 9, then the following may be needed
//    header('Cache-Control: max-age=1');
//
//// If you're serving to IE over SSL, then the following may be needed
//    header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
//    header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
//    header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
//    header('Pragma: public'); // HTTP/1.0
//
//    $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
//    $writer->save('php://output');