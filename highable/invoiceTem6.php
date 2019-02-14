<?php
require_once 'aidenfunc.php';
// modified by fa at 2019.01.15
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;

$inv =  $_SESSION['invoice'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/invoiceTem6.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
//$spreadsheet->getActiveSheet()->getDefaultColumnDimension()->setWidth(20);  //设置默认列宽
$spreadsheet->getActiveSheet()->getDefaultRowDimension()->setRowHeight(50); //行默认高度
$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(9);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(9);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(11);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('H')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('I')->setWidth(15);  //列宽度
$spreadsheet->getDefaultStyle()->getFont()->setSize(10);

$border = \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN;
$h_center = \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER;
$v_center = \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER;


$sheet = $spreadsheet->getActiveSheet();

for ($i=0;$i<100;$i++){
    $sheet->getRowDimension($i)->setRowHeight(15); //列高度
}

$sheet->getRowDimension(1)->setRowHeight(20); //列高度
$sheet->setCellValue("A1",$inv['remark']['poheader']['poheada1']);
setCell($sheet,'A2',$inv['remark']['poheader']['poheada2'].', '.$inv['remark']['poheader']['poheada3'],$noborderCenter);
$tel = $inv['remark']['poheader']['poheada4'].'  '.$inv['remark']['poheader']['poheada5'];
setCell($sheet,'A3',$tel,$noborderCenter);

//fill header
$sheet->setCellValue("F5", 'INVOICE NO.'.$inv['invoicedata']['invoiceNumber']);
$sheet->setCellValue("C6", $inv['tosb']);
$sheet->setCellValue("C7", $inv['invoicedata']['a1']);
$sheet->setCellValue("C8", $inv['invoicedata']['a2']);
$sheet->setCellValue("C9", 'Attn'.$inv['invoicedata']['a3']);
$sheet->setCellValue("J7",$inv['invoicedate']);

//fill bottom
$sheet->setCellValue('E29',$inv['remark']['bottomremark'][0]);
//$sheet->setCellValue('D22',$inv['remark']['bottomremark'][1]);
$sheet->setCellValue('E32',$inv['remark']['c1']);
$sheet->setCellValue('E33',$inv['remark']['c2']);
$sheet->setCellValue('E34',$inv['remark']['c3']);
$sheet->setCellValue('E35',$inv['remark']['c4']);
$sheet->setCellValue('E36',$inv['remark']['c5']);
setCell($sheet,'E37',$inv['remark']['c6'],$noborderboldfontLeft);
setCell($sheet,'F46',$inv['remark']['c7'],$noborderCenter);

//fill main content
{
    //form header
    {
        //three description input
        {
            $sheet->setCellValue("D14", $inv['invoiceform']['ba1'][0]);
            $sheet->setCellValue("D15", $inv['invoiceform']['ba1'][4]);
            $sheet->setCellValue("D16", $inv['invoiceform']['ba1'][5]);

            setCell($sheet,'E37',$inv['remark']['c6'],$noborderboldfontLeft);
        }

        //Unit Price , Ammount , Precent of ammount
        {
            $sheet->setCellValue("I13", $inv['invoiceform']['ba1'][1]);
            //$sheet->setCellValue("J13", $inv['invoiceform']['ba1'][2]);
            setCell($sheet,'J13',$inv['invoiceform']['ba1'][2],$noborderCenter);
            //$sheet->setCellValue("K13", $inv['invoiceform']['ba1'][3]);
            setCell($sheet,'K13',$inv['invoiceform']['ba1'][3].'%',$noborderCenter);
        }
    }

    //form footer
    {
        //total pcs and package
        $sheet->setCellValue("B21", $inv['invoiceform']['coltb']);
        $sheet->setCellValue("B23", $inv['invoiceform']['ba1'][7]);
        //total ammount
        $sheet->setCellValue("J23", $inv['invoiceform']['coltc']);
        //total precent of ammount
        $sheet->setCellValue("J21", $inv['invoiceform']['ba1'][6]);
        //remark
        $sheet->setCellValue("D25", $inv['invoiceform']['formremark'][2]);

        $sheet->setCellValue("E21", 'Less '.$inv['invoiceform']['ba1'][3].'% DOWN PAYMENT BEFORE SHIPMENT');


        $sheet->setCellValue("D19", 'OUR JOB NO.');
        $sheet->setCellValue("E19", $inv['invoiceform']['formremark'][0]);
        $sheet->setCellValue("F19", $inv['invoiceform']['formremark'][1]);
    }

    //form data
    {
        for ($i=$inv['invoiceform']['brrnum']-1,$j=$inv['invoiceform']['formnum']-1;$j>=0&&$i>=0;$j--,$i--){
            add_row($inv['invoiceform'],$i,$j);
        }
    }

}
function add_row($data,$i,$j){
    global $sheet;
    $sheet->insertNewRowBefore(18,1);

    //quantity
    $sheet->setCellValue("A18", "**");
    $sheet->setCellValue("B18", $data['b1'][$j]);
    $sheet->setCellValue("C18", "**PCS");
    //Po No.
    $sheet->setCellValue("D18", "PO No.:  ".$data['b4'][$j]);
    //Color
    $sheet->setCellValue("E18", "COLOUR:  ".$data['b5'][$j]);
    //our job No.
    //$sheet->setCellValue("F18", "OUR JOB NO.:  ".$data['b6'][$j]);
    $sheet->setCellValue("F18", $data['b6'][$j]);
    //description
    $sheet->setCellValue("G18", $data['b7'][$j]);
    //unit price
    $sheet->setCellValue("I18", $data['b8'][$j]);
    //amount
    $sheet->setCellValue("J18", $data['b9'][$j]);
    //precent of amount
    $sheet->setCellValue("K18", $data['b3'][$j]);
}

$spreadsheet->getActiveSheet()->getPageSetup()
    ->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);  //A4
$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页
//unset($_SESSION['invoice'] ); //注销SESSION

$filenameout = "Invoice_{$inv['shortname']}_{$inv['invoiceno']}";
outExcel($spreadsheet,$filenameout);