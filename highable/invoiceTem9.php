<?php

header("Content-type: text/html; charset=utf-8");
//KM  && NEXT

require_once 'aidenfunc.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;

$inv =  $_SESSION['invoice'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/invoiceTem9.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$spreadsheet->getActiveSheet()->setTitle("sheet1");
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
//$spreadsheet->getActiveSheet()->getDefaultColumnDimension()->setWidth(20);  //设置默认列宽
$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(16);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(16);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(16);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(16);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(16);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(16);  //列宽度
//$spreadsheet->getActiveSheet()->getColumnDimension('K')->setWidth(16);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('I')->setWidth(18);  //列宽度
//$spreadsheet->getActiveSheet()->getColumnDimension('J')->setWidth(16);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('L')->setWidth(20);  //列宽度
$spreadsheet->getDefaultStyle()->getFont()->setSize(7);



//填数据
//$sheet->mergeCells("J1:K1");
$sheet->setCellValue("A1",$inv['remark']['poheader']['poheada1']);
setCell($sheet,'A2',$inv['remark']['poheader']['poheada2'],$noborderCenter);
setCell($sheet,'A3',$inv['remark']['poheader']['poheada3'],$noborderCenter);
$tel = $inv['remark']['poheader']['poheada4'].'  '.$inv['remark']['poheader']['poheada5'];
setCell($sheet,'A4',$tel,$noborderCenter);
//setCell($sheet,'A6',$intem1['remark']['poheader']['poheada5'],$noborderCenter);


$sheet->setCellValue('A8', $inv["invoicedata"]["a13"]);
$spreadsheet->getActiveSheet()->setCellValue('E8', $inv["invoicedata"]["a14"]);
$spreadsheet->getActiveSheet()->setCellValue('I7', 'Invoice No.'.$inv["invoicedata"]["invoiceNumber"]);
$spreadsheet->getActiveSheet()->setCellValue('I13', 'Page: '.$inv["invoicedata"]["a1"]);

$spreadsheet->getActiveSheet()->setCellValue('J11', $inv["invoicedate"]);

$spreadsheet->getActiveSheet()->setCellValue('A18', $inv["invoicedata"]["a2"]);
$spreadsheet->getActiveSheet()->setCellValue('C18', $inv["invoicedata"]["a3"]);
$spreadsheet->getActiveSheet()->setCellValue('E18', $inv["invoicedata"]["a4"]);
$spreadsheet->getActiveSheet()->setCellValue('G18', $inv["invoicedata"]["a5"]);
//$spreadsheet->getActiveSheet()->setCellValue('I18', $inv["invoicedata"]["a6"]);
setMergeCells($sheet,'I18:J18','I18',$inv["invoicedata"]["a6"],$Size8noborderLeft);
//$spreadsheet->getActiveSheet()->setCellValue('K18', $inv["invoicedata"]["a7"]);
setMergeCells($sheet,'K18:L18','K18',$inv["invoicedata"]["a7"],$Size8noborderLeft);
//$spreadsheet->getActiveSheet()->setCellValue('A21', $inv["invoicedata"]["a8"]); //$Size8bordersLeft
setMergeCells($sheet,'A20:B21','A20',$inv["invoicedata"]["a8"],$Size8noborderLeft);

//$spreadsheet->getActiveSheet()->setCellValue('C21', $inv["invoicedata"]["a9"]);
setMergeCells($sheet,'C20:D21','C20',$inv["invoicedata"]["a9"],$Size8noborderLeft);
//$spreadsheet->getActiveSheet()->setCellValue('E21', $inv["invoicedata"]["a10"]);
setMergeCells($sheet,'E20:F21','E20',$inv["invoicedata"]["a10"],$Size8noborderLeft);

//$spreadsheet->getActiveSheet()->setCellValue('G21', $inv["invoicedata"]["a11"]);
setMergeCells($sheet,'G20:H21','G20',$inv["invoicedata"]["a11"],$Size8noborderLeft);

//$spreadsheet->getActiveSheet()->setCellValue('I21', $inv["invoicedata"]["a12"]);
setMergeCells($sheet,'I20:L21','I20',$inv["invoicedata"]["a12"],$Size8noborderLeft);

// 中间表格
$spreadsheet->getActiveSheet()->setCellValue('J24', '('.$inv["invoiceform"]["b12"].')');
$spreadsheet->getActiveSheet()->setCellValue('L24', '('.$inv["invoiceform"]["b13"].')');

$spreadsheet->getActiveSheet()->setCellValue('L32', $inv["invoiceform"]["coltb"]);
$spreadsheet->getActiveSheet()->setCellValue('C35', $inv["invoiceform"]["formremark"]);

// BOTTOM
$spreadsheet->getActiveSheet()->setCellValue('D40', $inv["remark"]["bottomremark"]["0"]);
$spreadsheet->getActiveSheet()->setCellValue('D42', $inv["remark"]["bottomremark"]["1"]);

//动态部分
$nowcol = 27;

foreach ($inv["invoiceform"]["b1"] as $item => $value) {
    if ($item > 0) {
        $spreadsheet->getActiveSheet()->insertNewRowBefore($nowcol, 4);
    }
    $spreadsheet->getActiveSheet()->setCellValue('A'.$nowcol, $value);
    $spreadsheet->getActiveSheet()->getStyle('A'.$nowcol)->applyFromArray($Size8noborderLeft);
    $nowcol += 4;
}

//$nowcol = 28;
////foreach ($inv["invoiceform"] as $item => $value) {
////    $formarr = array('A'.$nowcol,'B'.$nowcol,'D'.$nowcol,'E'.$nowcol,'I'.$nowcol,'J'.$nowcol,'L'.$nowcol);
////    for ($i = 2, $y = 0; $i <= 8; $i++, $y++) {
////        $spreadsheet->getActiveSheet()->setCellValue($formarr[$y], $value['b'.$i]);
////    }
////    $nowcol += 4;
////}
//
$nowcol = 28;
foreach ($inv["invoiceform"]["b2"] as $item => $value) {

    $spreadsheet->getActiveSheet()->setCellValue('A'.$nowcol, $value);
    $spreadsheet->getActiveSheet()->getStyle('A'.$nowcol)->applyFromArray($Size8noborderLeft);
    $nowcol += 4;
}

$nowcol = 28;
foreach ($inv["invoiceform"]["b3"] as $item => $value) {

    $spreadsheet->getActiveSheet()->setCellValue('B'.$nowcol, $value);
    $spreadsheet->getActiveSheet()->getStyle('B'.$nowcol)->applyFromArray($Size8noborderLeft);
    $nowcol += 4;
}
$nowcol = 28;
foreach ($inv["invoiceform"]["b4"] as $item => $value) {

    $spreadsheet->getActiveSheet()->setCellValue('D'.$nowcol, $value);
    $spreadsheet->getActiveSheet()->getStyle('D'.$nowcol)->applyFromArray($Size8noborderCenter);
    $nowcol += 4;
}

$nowcol = 28;
foreach ($inv["invoiceform"]["b5"] as $item => $value) {

//    $spreadsheet->getActiveSheet()->setCellValue('E'.$nowcol, $value);
//    $spreadsheet->getActiveSheet()->getStyle('E'.$nowcol)->applyFromArray($Size8noborderCenter);
    setMergeCells($sheet,"E{$nowcol}:G{$nowcol}",'E'.$nowcol,$value,$Size8noborderLeft);
    $nowcol += 4;
}
$nowcol = 28;
foreach ($inv["invoiceform"]["b6"] as $item => $value) {

    $spreadsheet->getActiveSheet()->setCellValue('I'.$nowcol, $value);
    $spreadsheet->getActiveSheet()->getStyle('I'.$nowcol)->applyFromArray($Size8noborderCenter);
    $nowcol += 4;
}
$nowcol = 28;
foreach ($inv["invoiceform"]["b7"] as $item => $value) {

    $spreadsheet->getActiveSheet()->setCellValue('J'.$nowcol, $value);
    $spreadsheet->getActiveSheet()->getStyle('J'.$nowcol)->applyFromArray($Size8noborderCenter);
    $nowcol += 4;
}
$nowcol = 28;
foreach ($inv["invoiceform"]["b8"] as $item => $value) {
    $spreadsheet->getActiveSheet()->setCellValue('L'.$nowcol, $value);
    $spreadsheet->getActiveSheet()->getStyle('L'.$nowcol)->applyFromArray($Size8noborderCenter);
    $nowcol += 4;
}


$nowcol = 29;
foreach ($inv["invoiceform"]["b9"] as $item => $value) {
    $spreadsheet->getActiveSheet()->setCellValue('A'.$nowcol, $value);
    $spreadsheet->getActiveSheet()->getStyle('A'.$nowcol)->applyFromArray($Size8noborderLeft);
    $nowcol += 4;
}

$nowcol = 30;
foreach ($inv["invoiceform"]["b10"] as $item => $value) {
//    if ($item > 1) {
//        $spreadsheet->getActiveSheet()->insertNewRowBefore($nowcol, 4);
//    }
    $spreadsheet->getActiveSheet()->setCellValue('E'.$nowcol, $value);
    $spreadsheet->getActiveSheet()->getStyle('E'.$nowcol)->applyFromArray($Size8noborderCenter);
    $nowcol += 4;
}

$spreadsheet->getActiveSheet()->getPageSetup()
    ->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);// A4
$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

//unset($_SESSION['invoice'] ); //注销SESSION

$filenameout = "Invoice_".$inv['shortname'];
outExcel($spreadsheet,$filenameout);