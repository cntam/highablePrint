<?php
session_start();
// modified by fa at 2019.01.16

header("Content-type: text/html; charset=utf-8");
require_once('autoloadconfig.php');  //判断是否在线
require_once ('img.php');

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;
$intem1 =  $_SESSION['invoiceTem7'];

$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/invoiceTem7.xlsx');
$sheet = $spreadsheet->getActiveSheet();
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
$spreadsheet->getDefaultStyle()->getFont()->setSize(10);
$sheet = $spreadsheet->getActiveSheet();

for ($i=0;$i<100;$i++){
    $sheet->getRowDimension($i)->setRowHeight(20); //列高度
}

//fill header
$sheet->setCellValue("A6", 'Invoice NO.'.$intem1['invoicedata']['invoiceNumber']);
$sheet->setCellValue("C8",  $intem1['tosb']);
$sheet->setCellValue("C9",  $intem1['invoicedata']['a1']);
$sheet->setCellValue("C10", $intem1['invoicedata']['a2']);
$sheet->setCellValue("C11", $intem1['invoicedata']['a3']);
$sheet->setCellValue("M8",$intem1['invoicedate']);

//fill bottom
$sheet->setCellValue('G18','COUNTRY OF ORIGIN:  '.$intem1['remark']['bottomremark'][0]);
$sheet->setCellValue('G20',$intem1['remark']['bottomremark'][1]);
$sheet->setCellValue('H25',$intem1['remark']['c1']);
$sheet->setCellValue('H26',$intem1['remark']['c2']);
$sheet->setCellValue('H27',$intem1['remark']['c3']);
$sheet->setCellValue('H28',$intem1['remark']['c4']);


//fill main content
{
    //form header
    {
        //Unit Price , total Ammount ,total carton
        {
            $sheet->setCellValue("L13", $intem1['invoiceform']['ba1'][0]);
            $sheet->setCellValue("M13", $intem1['invoiceform']['ba1'][1]);
            $sheet->setCellValue("M15", $intem1['invoiceform']['coltc']);
            $sheet->setCellValue("B15", $intem1['invoiceform']['coltb']);
            $sheet->setCellValue("C15", 'Carton');
        }
    }

    //form footer
    {
        //total pcs and package
        $sheet->setCellValue("G15", $intem1['invoiceform']['formremark']);
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
    $sheet->insertNewRowBefore(14, 1);

    //quantity
    $sheet->setCellValue("B14", $data['b1'][$j]);
    $sheet->setCellValue("C14", 'Carton');
    $sheet->setCellValue("D14", $data['b3'][$j]);
    $sheet->setCellValue("E14", '**mts');
    //description
    $sheet->setCellValue("G14", $data['b5'][$j]);
    //color
    $sheet->setCellValue("J14", $data['b6'][$j]);
    //color No.
    $sheet->setCellValue("K14", $data['b7'][$j]);
    //unit price
    $sheet->setCellValue("L14", $data['b8'][$j]);
    //amount
    $sheet->setCellValue("M14", $data['b9'][$j]);



}

$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页
//unset($_SESSION['invoiceTem7'] ); //注销SESSION
$output=  ($_GET['action'] == 'formdown' )? 1:0;
$nt = date("YmdHis",time()); //转换为日期。
$filenameout = 'intem1out'.$nt.'.xlsx';
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
