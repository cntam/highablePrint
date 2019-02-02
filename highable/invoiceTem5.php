<?php
require_once 'aidenfunc.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;

$inv =  $_SESSION['invoice'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/invoiceTem5.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle("sheet1");
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
//$spreadsheet->getActiveSheet()->getDefaultColumnDimension()->setWidth(20);  //设置默认列宽
//$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(16);  //列宽度
//$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(16);  //列宽度
//$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(16);  //列宽度
//$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(16);  //列宽度
$rowarr = range('E','M');
foreach ($rowarr as $value){
    $sheet->getColumnDimension($value)->setWidth(20);  //列宽度
}

$spreadsheet->getDefaultStyle()->getFont()->setSize(7);



//FILL SHEET HEADER
{
    $sheet->setCellValue("A1",$inv['remark']['poheader']['poheada1']);
    setCell($sheet,'A2',$inv['remark']['poheader']['poheada2'],$noborderCenter);
    setCell($sheet,'A3',$inv['remark']['poheader']['poheada3'],$noborderCenter);
    $tel = $inv['remark']['poheader']['poheada4'].'  '.$inv['remark']['poheader']['poheada5'];
    setCell($sheet,'A4',$tel,$noborderCenter);


    $sheet->setCellValue('A6', 'INVOICE NO.' . $inv["invoiceno"]);

    $sheet->setCellValue('C7', $inv["tosb"]);
    $sheet->setCellValue('C8', $inv["invoicedata"]["a1"]);
    $sheet->setCellValue('C9', $inv["invoicedata"]["a2"]);

    $sheet->setCellValue('L9', $inv["invoicedate"]);

    $sheet->setCellValue('G11', $inv["tosb"]);
    $sheet->setCellValue('G12', $inv["invoicedata"]["a3"]);
    $sheet->setCellValue('G13', $inv["invoicedata"]["a4"]);
    $sheet->setCellValue('G14', $inv["invoicedata"]["a5"]);
    $sheet->setCellValue('G15', $inv["invoicedata"]["a6"]);
    $sheet->setCellValue('G16', $inv["invoicedata"]["a7"]);

    $sheet->setCellValue('K16', $inv["invoiceform"]["ba1"][0]);
    $sheet->setCellValue('L16', $inv["invoiceform"]["ba1"][1]);
}

//$sheet->setCellValue('K16', $inv["invoicedata"]["a6"]);
//$sheet->setCellValue('L16', $inv["invoicedata"]["a6"]);
$sheet->setCellValue('M16', $inv["invoicedata"]["a8"].'%');

////中间表格固定内容
$sheet->setCellValue('A40', $inv["invoicedata"]["a9"]);
$sheet->setCellValue('B40', $inv["invoicedata"]["a10"]);
$sheet->setCellValue('C40', $inv["invoicedata"]["a11"]);
//setCell($sheet,'C40',$inv["invoicedata"]["a11"],$Size8noborderCenter);
$sheet->setCellValue('D40', $inv["invoicedata"]["a12"]);

//底部注释及银行信息
$sheet->setCellValue('G41', 'Less '.$inv["invoicedata"]["a8"].'%DOWN PAYMENT AND CQ COST  BEFORE SHIPMENT');

$sheet->setCellValue('L41', $inv["invoicedata"]["a14"]);
$sheet->setCellValue('L42', $inv["invoicedata"]["a15"]);

setMergeCells($sheet,'E44:J45','E44',$inv["invoiceform"]["formremark"],$Size8noborderLeft);

$sheet->setCellValue('E48', 'ORIGIN OF ORIGIN:'.$inv["remark"]["bottomremark"]["0"]);
$sheet->setCellValue('G50', $inv["remark"]["bottomremark"]["1"]);

$sheet->setCellValue('F53', $inv["remark"]["c1"]);
$sheet->setCellValue('F54', $inv["remark"]["c2"]);
$sheet->setCellValue('F55', $inv["remark"]["c3"]);
$sheet->setCellValue('F56', $inv["remark"]["c4"]);
setCell($sheet,'G58',$inv["remark"]["c5"],$Size8noborderCenter);

////中部表格动态
$row = 20;
foreach ($inv["invoiceform"]["b1"] as $item => $value) {
    if ($item > 4) {
        $sheet->insertNewRowBefore($row , 4);
    }
    $sheet->setCellValue('E'.$row, $value);
    $row += 4;
}
//21行
if ($inv["invoiceform"]["formnum"] > 0) {
    for ($a = 1; $a <= 13 ; $a++) {
        $row = 21;
        $col = chr(64 + $a); // A
        foreach ($inv["invoiceform"]['b'.($a + 1)] as $value) {
            $sheet->setCellValue($col.$row, $value);
            $row += 4;
        }
    }

}
//
//22行
if (count($inv["invoiceform"]["formnum"]) > 0) {
    for ($a = 1, $b = 15; $a <= 11  ; $a++, $b++) {
        $row = 22;
        $col = chr(66 + $a); // C
        foreach ($inv["invoiceform"]['b'.$b] as $value) {
            $sheet->setCellValue($col.$row, $value);
            $row += 4;
        }
    }

}

$spreadsheet->getActiveSheet()->getPageSetup()
    ->setOrientation(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE);  //横放置
$spreadsheet->getActiveSheet()->getPageSetup()
    ->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);  //A4

$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页
//unset($_SESSION['invoice'] ); //注销SESSION

$filenameout = "Invoice_{$inv['shortname']}_{$inv['invoiceno']}";
outExcel($spreadsheet,$filenameout);
