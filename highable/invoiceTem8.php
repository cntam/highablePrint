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
$intem1 =  $_SESSION['invoiceTem8'];

$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/invoiceTem8.xlsx');
$sheet = $spreadsheet->getActiveSheet();
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
$spreadsheet->getDefaultStyle()->getFont()->setSize(10);
$sheet = $spreadsheet->getActiveSheet();


//fill header
$sheet->setCellValue("A6", 'Invoice NO.'.$intem1['invoicedata']['invoiceNumber']);
$sheet->setCellValue("B8",  $intem1['tosb']);
$sheet->setCellValue("B9",  $intem1['invoicedata']['a1']);
$sheet->setCellValue("B10", $intem1['invoicedata']['a2']);
$sheet->setCellValue("B11", $intem1['invoicedata']['a3']);
$sheet->setCellValue("L8",$intem1['invoicedate']);

//fill main content
{
    //form header
    {
        //description
        {
            $sheet->setCellValue("F16", $intem1['invoiceform']['ba1'][0]);
            $sheet->setCellValue("F18", $intem1['invoiceform']['ba1'][3]);
        }
    }

    //form footer
    {
        //total ctins
        $sheet->setCellValue("A20", $intem1['invoiceform']['coltb']);
        //total amount
        $sheet->setCellValue("L20", $intem1['invoiceform']['coltc']);
        $sheet->setCellValue("B20", "CTINS");
        $sheet->setCellValue('F21',$intem1['remark']['bottomremark'][0]);
        $sheet->setCellValue('G23',$intem1['remark']['bottomremark'][1]);
        $sheet->setCellValue('G25',$intem1['remark']['bottomremark'][2]);
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
    $sheet->insertNewRowBefore(20, 4);

    //quantity
    $sheet->setCellValue("A20", $data['b1'][$j]);
    $sheet->setCellValue("B20", 'CTINS');
    $sheet->setCellValue("C20", $data['b3'][$j]);
    $sheet->setCellValue("E20", 'PCS');

    $sheet->setCellValue("I20", $data['b7'][$j]);

    $sheet->setCellValue("F20", 'STYLE NO.:');
    $sheet->setCellValue("G20", $data['b4'][$j]);
    $sheet->setCellValue("F21", 'STYLE CODE:');
    $sheet->setCellValue("G21", $data['b5'][$j]);
    $sheet->setCellValue("F22", 'ORDER NO.:');
    $sheet->setCellValue("G22", $data['b6'][$j]);

    //unit price amount
    $sheet->setCellValue("K20", $data['b8'][$j]);
    $sheet->setCellValue("L20", $data['b9'][$j]);

}

$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页
unset($_SESSION['invoiceTem8'] ); //注销SESSION
$output=  ($_GET['action'] == 'formdown' )? 1:0;
$nt = date("md",time()); //转换为日期。
$filenameout = 'Invoice_JIGSAW_'.$nt.'.xlsx';
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

