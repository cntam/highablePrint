<?php

header("Content-type: text/html; charset=utf-8");
require_once 'aidenfunc.php';
// modified by fa at 2019.01.16
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;

$intem1 =  $_SESSION['invoice'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/invoiceTem3.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
$spreadsheet->getActiveSheet()->getDefaultRowDimension()->setRowHeight(50); //行默认高度
$spreadsheet->getActiveSheet()->getColumnDimension('H')->setWidth(20);  //列宽度
$spreadsheet->getDefaultStyle()->getFont()->setSize(10);
$sheet = $spreadsheet->getActiveSheet();

for ($i=0;$i<100;$i++){
    $sheet->getRowDimension($i)->setRowHeight(15); //列高度
}
//
////fill header
$sheet->setCellValue("F8", 'INVOICE NO.'.$intem1['invoicedata']['invoiceNumber']);
$sheet->setCellValue("C9", $intem1['tosb']);
$sheet->setCellValue("C10", $intem1['invoicedata']['a1']);
$sheet->setCellValue("C11", $intem1['invoicedata']['a2']);
$sheet->setCellValue("C12", $intem1['invoicedata']['a3']);
$sheet->setCellValue("C13", $intem1['invoicedata']['a4']);
$sheet->setCellValue("J10",$intem1['invoicedate']);

//fill main content
{
    //form header
    {
        //four description input
        {
            $sheet->setCellValue("D17", $intem1['invoiceform']['ba1'][0]);
            $sheet->setCellValue("D19", $intem1['invoiceform']['ba1'][3]);
            $sheet->setCellValue("H19", $intem1['invoiceform']['ba1'][4]);
            $sheet->setCellValue("D21", $intem1['invoiceform']['ba1'][5]);
        }

        //Unit Price , Ammount
        {
            $sheet->setCellValue("I17", $intem1['invoiceform']['ba1'][1]);
            $sheet->setCellValue("J17", $intem1['invoiceform']['ba1'][2]);

        }
    }

    //form footer
    {
        //amount total
        $sheet->setCellValue("J31", $intem1['invoiceform']['coltc']);
        //package
        $sheet->setCellValue("B31", $intem1['invoiceform']['b13']);

        $sheet->setCellValue('D33', $intem1['remark']['bottomremark'][0]);
        $sheet->setCellValue('D34', $intem1['remark']['bottomremark'][1]);
        $sheet->setCellValue('E36', $intem1['remark']['c1']);
        $sheet->setCellValue('E37', $intem1['remark']['c2']);
        $sheet->setCellValue('E38', $intem1['remark']['c3']);
        $sheet->setCellValue('E39', $intem1['remark']['c4']);
        $sheet->setCellValue('E40', $intem1['remark']['c5']);
        $sheet->setCellValue('E41', $intem1['remark']['c6']);
        $sheet->setCellValue('D42', $intem1['remark']['c7']);
    }
    //form data
    {
        for ($i=$intem1['invoiceform']['brrnum']-1,$j=$intem1['invoiceform']['formnum']-1;$j>=0&&$i>=0;$j--,$i--){
            add_row($intem1['invoiceform'],$i,$j);
        }
    }

}
function add_row($data,$i,$j)
{
    global $sheet;
    $sheet->insertNewRowBefore(26,5);

    $sheet->setCellValue("D26", $data['b3'][$j]);
    //quantity
    $sheet->setCellValue("A27", "**");
    $sheet->setCellValue("B27", $data['b1'][$j]);
    $sheet->setCellValue("C27", "**PCS");
    //Po No.
    $sheet->setCellValue("D27", "PO No.:  ");
    $sheet->setCellValue("E27", $data['b5'][$j]);
    //Color
    $sheet->setCellValue("D28", "COLOUR:  ");
    $sheet->setCellValue("E28", $data['b7'][$j]);
    //our job No.
    $sheet->setCellValue("D29", "OUR JOB NO.:  ");
    $sheet->setCellValue("E29", $data['b9'][$j]);

    $sheet->setCellValue("G29", $data['b10'][$j]);

    //unit price
    $sheet->setCellValue("I26", $data['b11'][$j]);
    //amount
    $sheet->setCellValue("J26", $data['b12'][$j]);


}



$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页
//unset($_SESSION['invoiceTem3'] ); //注销SESSION
$output=  ($_GET['action'] == 'formdown' )? 1:0;
$nt = date("md",time()); //转换为日期。
$filenameout = "Invoice_{$invoiceTem4['shortname']}_".$nt.'.xlsx';
if($output){
    // Redirect output to a client’s web browser (Xlsx)
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment;filename='."$filenameout");
    header('Cache-Control: max-age=0');
// If you're serving to IE 9, then the following may be needed
    header('Cache-Control: max-age=1');

// If you're serving to IE over SSL, then the following may be needed
    header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
    header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
    header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
    header('Pragma: public'); // HTTP/1.0

    $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
    $writer->save('php://output');
}else{
    $writer = new Xlsx($spreadsheet);
    $writer->save('../output/'.$filenameout);

    $FILEURL = 'http://allinone321.com/highable/output/'.$filenameout;
    $MSFILEURL = 'http://view.officeapps.live.com/op/view.aspx?src='. urlencode($FILEURL);
    //echo "<a href= 'http://view.officeapps.live.com/op/view.aspx?src=". urlencode($FILEURL)."' target='_blank' >跳轉--{$filename}</a>";
    Header("Location:{$MSFILEURL}");
};

