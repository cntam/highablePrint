<?php
header("Content-type: text/html; charset=utf-8");
/*港源行國際有限公司*/

require_once 'aidenfunc.php';


$packinglistTem3 =  $_SESSION['packinglist'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/packinglistTem3.xlsx');
$spreadsheet->getActiveSheet()->setTitle("sheet1");
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

////form static size
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
//
$sheet->setCellValue('J32', $packinglistTem3["invoicedata"]["a21"]);
$sheet->setCellValue('K32', $packinglistTem3["invoicedata"]["a27"]);
//
////form 动态
////COLOUR & SIZE BREAKDOWN 第一行
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

unset($_SESSION['packinglist'] ); //注销SESSION

$filenameout = "Packinglist_".$packinglistTem3['shortname'];
outExcel($spreadsheet,$filenameout);
