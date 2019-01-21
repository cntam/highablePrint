<!--// Modified by 俊伟-->
<?php
session_start();
header("Content-type: text/html; charset=utf-8");
/*港源行國際有限公司*/

require_once('autoloadconfig.php');  //判断是否在线

require_once ('img.php');

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;

$packinglistTem3 =  $_SESSION['packinglist'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/packinglistTem3.xlsx');

$sheet = $spreadsheet->getActiveSheet();

// 填数据
//header
$sheet->setCellValue('L8', $packinglistTem3["invoicedata"]["a28"]);
$sheet->setCellValue('A10', 'INVOICE NO. '.$packinglistTem3["invoicedata"]["invoiceNumber"]);
$sheet->setCellValue('L10', $packinglistTem3["invoicedata"]["a29"]);

//form static
$sheet->setCellValue('A13', $packinglistTem3["invoicedata"]["a1"]);
$sheet->setCellValue('A14', $packinglistTem3["invoicedata"]["a4"]);
$sheet->setCellValue('A15', $packinglistTem3["invoicedata"]["a6"]);
$sheet->setCellValue('A16', $packinglistTem3["invoicedata"]["a8"]);
$sheet->setCellValue('A17', $packinglistTem3["invoicedata"]["a10"]);
$sheet->setCellValue('A18', $packinglistTem3["invoicedata"]["a11"]);

$sheet->setCellValue('D13', $packinglistTem3["invoicedata"]["a2"]);
$sheet->setCellValue('D14', $packinglistTem3["invoicedata"]["a5"]);
$sheet->setCellValue('D15', $packinglistTem3["invoicedata"]["a7"]);
$sheet->setCellValue('D16', $packinglistTem3["invoicedata"]["a9"]);


$sheet->setCellValue('E13', $packinglistTem3["invoicedata"]["a3"]);

//form static size
$sheet->setCellValue('E18', $packinglistTem3["invoicedata"]["a12"]);
$sheet->setCellValue('F18', $packinglistTem3["invoicedata"]["a13"]);
$sheet->setCellValue('G18', $packinglistTem3["invoicedata"]["a14"]);
$sheet->setCellValue('H18', $packinglistTem3["invoicedata"]["a15"]);
$sheet->setCellValue('I18', $packinglistTem3["invoicedata"]["a16"]);
$sheet->setCellValue('J18', $packinglistTem3["invoicedata"]["a17"]);
$sheet->setCellValue('K18', $packinglistTem3["invoicedata"]["a18"]);

$sheet->setCellValue('C21', $packinglistTem3["invoicedata"]["a21"]);
$sheet->setCellValue('D21', $packinglistTem3["invoicedata"]["a22"]);
$sheet->setCellValue('I21', $packinglistTem3["invoicedata"]["a23"]);
$sheet->setCellValue('L21', $packinglistTem3["invoicedata"]["a24"]);
$sheet->setCellValue('M21', $packinglistTem3["invoicedata"]["a25"]);

$sheet->setCellValue('J32', $packinglistTem3["invoicedata"]["a21"]);
$sheet->setCellValue('K32', $packinglistTem3["invoicedata"]["a27"]);

//form 动态
//COLOUR & SIZE BREAKDOWN 第一行
if ($packinglistTem3["invoiceform"]["brownum"] > 0) {
    $arr = array('D', 'E', 'I', 'J', 'K');
    for ($a = 0, $b = 13; $a < count($arr) ; $a++, $b++) {
        $row = 25;
        $col = $arr[$a];
        foreach ($packinglistTem3["invoiceform"]['b'.$b] as $item => $value) {
            if (($item > 0)&&($b == 13)) {
                $sheet->insertNewRowBefore($row, 7);
                for ($x = 0; $x <= 6; $x++) {
                    $newRow = $row + $x;
                    $contextArrE = 'E'.$newRow;
                    $contextArrH = 'H'.$newRow;
                    $sheet->mergeCells("$contextArrE:$contextArrH");
                }
            }
            $sheet->setCellValue($col.$row, $value);
            $row += 7;
        }
    }
}
//COLOUR & SIZE BREAKDOWN 剩下7行
if ($packinglistTem3["invoiceform"]["brownum"] > 0) {
    $arr = array('D', 'E', 'I', 'J', 'K');
    for ($a = 0, $b = 18; $a < count($arr) ; $a++, $b++) {
        $row = 26;
        $col = $arr[$a];
        foreach ($packinglistTem3["invoiceform"]['b'.$b] as $item => $value) {
            $sheet->setCellValue($col.$row, $value);
            $row += 7;
        }
    }
    for ($a = 0, $b = 23; $a < count($arr) ; $a++, $b++) {
        $row = 27;
        $col = $arr[$a];
        foreach ($packinglistTem3["invoiceform"]['b'.$b] as $item => $value) {
            $sheet->setCellValue($col.$row, $value);
            $row += 7;
        }
    }
    for ($a = 0, $b = 28; $a < count($arr) ; $a++, $b++) {
        $row = 28;
        $col = $arr[$a];
        foreach ($packinglistTem3["invoiceform"]['b'.$b] as $item => $value) {
            $sheet->setCellValue($col.$row, $value);
            $row += 7;
        }
    }
    for ($a = 0, $b = 33; $a < count($arr) ; $a++, $b++) {
        $row = 29;
        $col = $arr[$a];
        foreach ($packinglistTem3["invoiceform"]['b'.$b] as $item => $value) {
            $sheet->setCellValue($col.$row, $value);
            $row += 7;
        }
    }
    for ($a = 0, $b = 38; $a < count($arr) ; $a++, $b++) {
        $row = 30;
        $col = $arr[$a];
        foreach ($packinglistTem3["invoiceform"]['b'.$b] as $item => $value) {
            $sheet->setCellValue($col.$row, $value);
            $row += 7;
        }
    }
    for ($a = 0, $b = 43; $a < count($arr) ; $a++, $b++) {
        $row = 31;
        $col = $arr[$a];
        foreach ($packinglistTem3["invoiceform"]['b'.$b] as $item => $value) {
            $sheet->setCellValue($col.$row, $value);
            $row += 7;
        }
    }
}

if ($packinglistTem3["invoiceform"]["brownum"] > 0) {
    for ($a = 1; $a <= 12 ; $a++) {
        $row = 19;
        $col = chr(65 + $a); // B
        foreach ($packinglistTem3["invoiceform"]['b'.$a] as $item => $value) {
            if (($item > 0)&&($a == 1)) {
                $sheet->insertNewRowBefore($row, 1);
            }
            $sheet->setCellValue($col.$row, $value);
            $row++;
        }
    }
}

$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

//unset($_SESSION['packinglist'] ); //注销SESSION

$output=  ($_GET['action'] == 'formdown' )? 1:0;
$nt = date("md",time()); //转换为日期。
$filenameout = 'Packinglist_LAUK_'.$nt.'.xlsx';
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
