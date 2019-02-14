<?php
header("Content-type: text/html; charset=utf-8");
/*港源行國際有限公司*/

require_once 'aidenfunc.php';


$pl =  $_SESSION['packinglist'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/packinglistTem3.xlsx');
$spreadsheet->getActiveSheet()->setTitle("sheet1");
$sheet = $spreadsheet->getActiveSheet();

// 填数据
$sheet->setCellValue("A1",$pl['remark']['poheader']['poheada1']);
setCell($sheet,'A2',$pl['remark']['poheader']['poheada2'],$noborderCenter);
setCell($sheet,'A3',$pl['remark']['poheader']['poheada3'],$noborderCenter);
setCell($sheet,'A4',$pl['remark']['poheader']['poheada4'],$noborderCenter);
setCell($sheet,'A5',$pl['remark']['poheader']['poheada5'],$noborderCenter);


//header
$sheet->setCellValue('L8', $pl["invoicedata"]["a28"]);
$sheet->setCellValue('A10', 'INVOICE NO. '.$pl["invoicedata"]["invoiceNumber"]);
$sheet->setCellValue('L10', $pl["invoicedata"]["a29"]);

//form static
$sheet->setCellValue('A13', $pl["invoicedata"]["a1"]);
$sheet->setCellValue('A14', $pl["invoicedata"]["a4"]);
$sheet->setCellValue('A15', $pl["invoicedata"]["a6"]);
$sheet->setCellValue('A16', $pl["invoicedata"]["a8"]);
$sheet->setCellValue('A17', $pl["invoicedata"]["a10"]);
$sheet->setCellValue('A18', $pl["invoicedata"]["a11"]);

$sheet->setCellValue('D13', $pl["invoicedata"]["a2"]);
$sheet->setCellValue('D14', $pl["invoicedata"]["a5"]);
$sheet->setCellValue('D15', $pl["invoicedata"]["a7"]);
$sheet->setCellValue('D16', $pl["invoicedata"]["a9"]);


$sheet->setCellValue('E13', $pl["invoicedata"]["a3"]);

////form static size
$sheet->setCellValue('E18', $pl["invoicedata"]["a12"]);
$sheet->setCellValue('F18', $pl["invoicedata"]["a13"]);
$sheet->setCellValue('G18', $pl["invoicedata"]["a14"]);
$sheet->setCellValue('H18', $pl["invoicedata"]["a15"]);
$sheet->setCellValue('I18', $pl["invoicedata"]["a16"]);
$sheet->setCellValue('J18', $pl["invoicedata"]["a17"]);
$sheet->setCellValue('K18', $pl["invoicedata"]["a18"]);

$sheet->setCellValue('C21', $pl["invoicedata"]["a21"]);
$sheet->setCellValue('D21', $pl["invoicedata"]["a22"]);
$sheet->setCellValue('I21', $pl["invoicedata"]["a23"]);
$sheet->setCellValue('L21', $pl["invoicedata"]["a24"]);
$sheet->setCellValue('M21', $pl["invoicedata"]["a25"]);
//
$sheet->setCellValue('J32', $pl["invoicedata"]["a21"]);
$sheet->setCellValue('K32', $pl["invoicedata"]["a27"]);
//
////form 动态
////COLOUR & SIZE BREAKDOWN 第一行
if ($pl["invoiceform"]["brownum"] > 0) {
    $arr = array('D', 'E', 'I', 'J', 'K');
    for ($a = 0, $b = 13; $a < count($arr) ; $a++, $b++) {
        $row = 25;
        $col = $arr[$a];
        foreach ($pl["invoiceform"]['b'.$b] as $item => $value) {
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
if ($pl["invoiceform"]["brownum"] > 0) {
    $arr = array('D', 'E', 'I', 'J', 'K');
    for ($a = 0, $b = 18; $a < count($arr) ; $a++, $b++) {
        $row = 26;
        $col = $arr[$a];
        foreach ($pl["invoiceform"]['b'.$b] as $item => $value) {
            $sheet->setCellValue($col.$row, $value);
            $row += 7;
        }
    }
    for ($a = 0, $b = 23; $a < count($arr) ; $a++, $b++) {
        $row = 27;
        $col = $arr[$a];
        foreach ($pl["invoiceform"]['b'.$b] as $item => $value) {
            $sheet->setCellValue($col.$row, $value);
            $row += 7;
        }
    }
    for ($a = 0, $b = 28; $a < count($arr) ; $a++, $b++) {
        $row = 28;
        $col = $arr[$a];
        foreach ($pl["invoiceform"]['b'.$b] as $item => $value) {
            $sheet->setCellValue($col.$row, $value);
            $row += 7;
        }
    }
    for ($a = 0, $b = 33; $a < count($arr) ; $a++, $b++) {
        $row = 29;
        $col = $arr[$a];
        foreach ($pl["invoiceform"]['b'.$b] as $item => $value) {
            $sheet->setCellValue($col.$row, $value);
            $row += 7;
        }
    }
    for ($a = 0, $b = 38; $a < count($arr) ; $a++, $b++) {
        $row = 30;
        $col = $arr[$a];
        foreach ($pl["invoiceform"]['b'.$b] as $item => $value) {
            $sheet->setCellValue($col.$row, $value);
            $row += 7;
        }
    }
    for ($a = 0, $b = 43; $a < count($arr) ; $a++, $b++) {
        $row = 31;
        $col = $arr[$a];
        foreach ($pl["invoiceform"]['b'.$b] as $item => $value) {
            $sheet->setCellValue($col.$row, $value);
            $row += 7;
        }
    }
}

if ($pl["invoiceform"]["brownum"] > 0) {
    for ($a = 1; $a <= 12 ; $a++) {
        $row = 19;
        $col = chr(65 + $a); // B
        foreach ($pl["invoiceform"]['b'.$a] as $item => $value) {
            if (($item > 0)&&($a == 1)) {
                $sheet->insertNewRowBefore($row, 1);
            }
            $sheet->setCellValue($col.$row, $value);
            $row++;
        }
    }
}
$spreadsheet->getActiveSheet()->getPageSetup()
    ->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);  //A4
$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

unset($_SESSION['packinglist'] ); //注销SESSION

$filenameout = "Packinglist_".$pl['shortname'];
outExcel($spreadsheet,$filenameout);
