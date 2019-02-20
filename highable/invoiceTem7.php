<?php
require_once 'aidenfunc.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;
$inv =  $_SESSION['invoice'];

$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/invoiceTem7.xlsx');
$sheet = $spreadsheet->getActiveSheet();
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
$spreadsheet->getDefaultStyle()->getFont()->setSize(10);
$sheet = $spreadsheet->getActiveSheet();

for ($i=2;$i<100;$i++){
    $sheet->getRowDimension($i)->setRowHeight(20); //列高度
}

$sheet->setCellValue("A1",$inv['remark']['poheader']['poheada1']);
setCell($sheet,'A2',$inv['remark']['poheader']['poheada2'],$noborderCenter);
setCell($sheet,'A3',$inv['remark']['poheader']['poheada3'],$noborderCenter);
$tel = $inv['remark']['poheader']['poheada4'];
setCell($sheet,'A4',$tel,$noborderCenter);
//setCell($sheet,'A6',$inv['remark']['poheader']['poheada5'],$noborderCenter);

//fill header
$sheet->setCellValue("A6", 'Invoice NO.'.$inv['invoicedata']['invoiceNumber']);
$sheet->setCellValue("B8",  $inv['tosb']);
$sheet->setCellValue("B9",  $inv['invoicedata']['a1']);
$sheet->setCellValue("B10", $inv['invoicedata']['a2']);
$sheet->setCellValue("B11", $inv['invoicedata']['a3']);
$sheet->setCellValue("L8",$inv['invoicedate']);

//fill bottom
$sheet->setCellValue('F18','COUNTRY OF ORIGIN:  '.$inv['remark']['bottomremark'][0]);
setCell($sheet,'F22','Remark:',$noborderLeft);
setCell($sheet,'G22',$inv['remark']['bottomremark'][1],$noborderLeft);

$sheet->setCellValue('G24',$inv['remark']['c1']);
$sheet->setCellValue('G25',$inv['remark']['c2']);
$sheet->setCellValue('G26',$inv['remark']['c3']);
$sheet->setCellValue('G27',$inv['remark']['c4']);
setCell($sheet,'G28',$inv['remark']['c5'],$noborderLeft);

//fill main content
{
    //form header
    {
        //Unit Price , total Ammount ,total carton
        {
            $sheet->setCellValue("L13", $inv['invoiceform']['ba1'][0]);
            $sheet->setCellValue("L13", $inv['invoiceform']['ba1'][1]);
            $sheet->setCellValue("L15", $inv['invoiceform']['coltc']);
            $sheet->setCellValue("B15", $inv['invoiceform']['coltb']);
            $sheet->setCellValue("C15", 'Carton');
        }
    }

    //form footer
    {
        //total pcs and package
        $sheet->setCellValue("F19", $inv['invoiceform']['formremark']);
    }

    //form data
    {
        for ($i=$inv['invoiceform']['brrnum']-1,$j=$inv['invoiceform']['formnum']-1;$j>=0&&$i>=0;$j--,$i--){
            add_row($inv['invoiceform'],$i,$j,$noborderLeft);
        }
    }

}

function add_row($data,$i,$j,$sheetstyle)
{
    global $sheet;
    $sheet->insertNewRowBefore(14, 1);

    //quantity
    $sheet->setCellValue("B14", $data['b1'][$j]);
    $sheet->setCellValue("C14", 'Carton');
    $sheet->setCellValue("D14", $data['b3'][$j]);
    $sheet->setCellValue("E14", '**mts');
    //description
    //$sheet->setCellValue("G14", $data['b5'][$j]);
    setMergeCells($sheet,'F14:H14','F14',$data['b5'][$j],$sheetstyle);
    //color
    $sheet->setCellValue("I14", $data['b6'][$j]);
    //color No.
    $sheet->setCellValue("J14", $data['b7'][$j]);
    //unit price
    $sheet->setCellValue("K14", $data['b8'][$j]);
    //amount
    $sheet->setCellValue("L14", $data['b9'][$j]);

}

$spreadsheet->getActiveSheet()->getPageSetup()
    ->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);  //A4
$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

unset($_SESSION['invoice'] ); //注销SESSION

$filenameout = "Invoice_".$inv['shortname'];
outExcel($spreadsheet,$filenameout);

