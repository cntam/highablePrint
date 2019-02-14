<?php
require_once 'aidenfunc.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;
$inv =  $_SESSION['invoice'];

$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/invoiceTem8.xlsx');
$sheet = $spreadsheet->getActiveSheet();
$spreadsheet->getActiveSheet()->setTitle("sheet1");
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
$spreadsheet->getDefaultStyle()->getFont()->setSize(10);
$sheet = $spreadsheet->getActiveSheet();

$spreadsheet->getActiveSheet()->getColumnDimension('J')->setWidth(20);  //列宽度


$sheet->setCellValue("A1",$inv['remark']['poheader']['poheada1']);
setCell($sheet,'A2',$inv['remark']['poheader']['poheada2'],$noborderCenter);
setCell($sheet,'A3',$inv['remark']['poheader']['poheada3'],$noborderCenter);
$tel = $inv['remark']['poheader']['poheada4'].'  '.$inv['remark']['poheader']['poheada5'];
setCell($sheet,'A4',$tel,$noborderCenter);

//fill header
$sheet->setCellValue("A6", 'Invoice NO.'.$inv['invoicedata']['invoiceNumber']);
$sheet->setCellValue("B8",  $inv['tosb']);
$sheet->setCellValue("B9",  $inv['invoicedata']['a1']);
$sheet->setCellValue("B10", $inv['invoicedata']['a2']);
$sheet->setCellValue("B11", $inv['invoicedata']['a3']);
$sheet->setCellValue("L8",$inv['invoicedate']);

//fill main content
{
    //form header
    {
        //description
        {
            $sheet->setCellValue("F16", $inv['invoiceform']['ba1'][0]);
            $sheet->setCellValue("F18", $inv['invoiceform']['ba1'][3]);
        }
    }

    //form footer
    {
        //total ctins
        setCell($sheet,'A20','**',$noborderCenter);
        setCell($sheet,'B20',$inv['invoiceform']['coltb'],$noborderCenter);

        //total amount
        $sheet->setCellValue("L20", $inv['invoiceform']['coltc']);
        $sheet->setCellValue("C20", "CTINS");

        //setCell($sheet,'F21',$inv["invoiceform"]["formremark"],$Size8noborderLeft);
        setMergeCells($sheet,'F21:J22','F21',$inv["invoiceform"]["formremark"],$Size8noborderLeft);

        $sheet->setCellValue('F24',$inv['remark']['bottomremark'][0]);
        $sheet->setCellValue('G26',$inv['remark']['bottomremark'][1]);
        $sheet->setCellValue('G28',$inv['remark']['bottomremark'][2]);


        $sheet->setCellValue('H30',$inv["remark"]["c1"]);
    }

    //form data
    {
        for ($i=$inv['invoiceform']['brrnum']-1,$j=$inv['invoiceform']['formnum']-1;$j>=0&&$i>=0;$j--,$i--){
            add_row($inv['invoiceform'],$i,$j,$noborderCenter);
        }
    }

}

function add_row($data,$i,$j,$exstyle=false)
{
    global $sheet;
    $sheet->insertNewRowBefore(20, 4);

    //quantity
    //$sheet->setCellValue("A20", $data['b1'][$j]);
    setCell($sheet,'A20','**',$exstyle);
    setCell($sheet,'B20',$data['b1'][$j],$exstyle);
    $sheet->setCellValue("C20", 'CTINS');

    //$sheet->setCellValue("C20", $data['b3'][$j]);
    setCell($sheet,'D20',$data['b3'][$j],$exstyle);
    $sheet->setCellValue("E20", 'PCS');

    //$sheet->setCellValue("I20", $data['b7'][$j]);
    setCell($sheet,'H20',$data['b7'][$j],$exstyle);
    setCell($sheet,'I20',$data['b11'][$j],$exstyle);
    setCell($sheet,'J20',$data['b12'][$j],$exstyle);

    $sheet->setCellValue("F20", 'STYLE NO.:');
    //$sheet->setCellValue("G20", $data['b4'][$j]);
    setCell($sheet,'G20',$data['b4'][$j],$exstyle);

    $sheet->setCellValue("F21", 'STYLE CODE:');
    //$sheet->setCellValue("G21", $data['b5'][$j]);
    setCell($sheet,'G21',$data['b5'][$j],$exstyle);

    $sheet->setCellValue("F22", 'ORDER NO.:');
    //$sheet->setCellValue("G22", $data['b6'][$j]);
    setCell($sheet,'G22',$data['b6'][$j],$exstyle);

    //unit price amount
    $sheet->setCellValue("K20", $data['b8'][$j]);
    $sheet->setCellValue("L20", $data['b9'][$j]);

}

$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页
unset($_SESSION['invoice'] ); //注销SESSION

$filenameout = "Invoice_".$inv['shortname'];
outExcel($spreadsheet,$filenameout);

