<?php
require_once 'aidenfunc.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;

$pl =  $_SESSION['packinglist'];

//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/packinglistTem6.xlsx');
$sheet = $spreadsheet->getActiveSheet();
$spreadsheet->getActiveSheet()->setTitle("sheet1");

//填数据
//header
//$sheet->setCellValue('P8', $pl["invoicedata"]["invoiceNumber"]);

$sheet->getRowDimension(2)->setRowHeight(20); //列高度
$sheet->setCellValue("A2",$pl['remark']['poheader']['poheada1']);
setCell($sheet,'A4',$pl['remark']['poheader']['poheada2'].' '.$pl['remark']['poheader']['poheada3'],$noborderCenter);
$tel = $pl['remark']['poheader']['poheada4'].'  '.$pl['remark']['poheader']['poheada5'];
setCell($sheet,'A5',$tel,$noborderCenter);

setCell($sheet,'P9',$pl['invoicedata']["invoicedate"],$noborderCenter);
$sheet->setCellValue('B10', $pl["invoicedata"]["a1"]);
$sheet->setCellValue('B11', $pl["invoicedata"]["a2"]);
$sheet->setCellValue('B12', $pl["invoicedata"]["a3"]);
$sheet->setCellValue('B13', 'Attn:'.$pl["invoicedata"]["a4"]);

$sheet->setCellValue('A19', $pl["invoicedata"]["a5"]);
$sheet->setCellValue('F19', $pl["invoicedata"]["a6"]);
$sheet->setCellValue('A20', $pl["invoicedata"]["a7"]);
//Size Breakdown
$sheet->setCellValue('D23', $pl["invoiceform"]["b1"]["0"]);
$sheet->setCellValue('E23', $pl["invoiceform"]["b1"]["1"]);
$sheet->setCellValue('F23', $pl["invoiceform"]["b1"]["2"]);
$sheet->setCellValue('G23', $pl["invoiceform"]["b1"]["3"]);
$sheet->setCellValue('H23', $pl["invoiceform"]["b1"]["4"]);
$sheet->setCellValue('I23', $pl["invoiceform"]["b1"]["5"]);
$sheet->setCellValue('J23', $pl["invoiceform"]["b1"]["6"]);
$sheet->setCellValue('K23', $pl["invoiceform"]["b1"]["7"]);
//COLOUR/SIZE
$sheet->setCellValue('D36', $pl["invoiceform"]["b1"]["8"]);
$sheet->setCellValue('E36', $pl["invoiceform"]["b1"]["9"]);
$sheet->setCellValue('F36', $pl["invoiceform"]["b1"]["10"]);
$sheet->setCellValue('G36', $pl["invoiceform"]["b1"]["11"]);
$sheet->setCellValue('H36', $pl["invoiceform"]["b1"]["12"]);
$sheet->setCellValue('I36', $pl["invoiceform"]["b1"]["13"]);
$sheet->setCellValue('J36', $pl["invoiceform"]["b1"]["14"]);
$sheet->setCellValue('K36', $pl["invoiceform"]["b1"]["15"]);


//form total
$sheet->setCellValue('L29', $pl["invoiceform"]["ba1"][1]);
$sheet->setCellValue('N29', $pl["invoiceform"]["ba1"][2]);
$sheet->setCellValue('O29', $pl["invoiceform"]["ba1"][3]);

$sheet->setCellValue('B31', $pl["invoiceform"]["ba1"][0]); //$inarr["invoiceform"]["ba1"][0] ;

//footer
//$sheet->setCellValue('C31', $pl["invoicedata"]["invoiceNumber"]);
$sheet->setCellValue('C42', $pl["invoiceform"]["ba1"][2]);
$sheet->setCellValue('C43', $pl["invoiceform"]["ba1"][3]);

//form动态
if ($pl["invoiceform"]["brownum"] > 0) {
//    COLOUR/SIZE
    for ($a = 1, $b = 18; $a <= 9 ; $a++, $b++) {
        $row = 37;
        $col = chr(67 + $a); // D
        foreach ($pl["invoiceform"]['b'.$b] as $item => $value) {
            if (($item > 0)&&($b == 18)) {
                $sheet->insertNewRowBefore($row, 1);
            }
            $sheet->setCellValue($col.$row, $value);
            $row++;
        }
    }
    for ($a = 0; $a < $pl["invoiceform"]["brownum"] ; $a++) {
        $row = 37;
        foreach ($pl["invoiceform"]["b17"] as $item => $value) {
            $sheet->setCellValue('B'.$row, $value);
            $row++;
        }
    }
//    C/NO
    for ($a = 1, $b = 2; $a <= 12  ; $a++, $b++) {
        $row = 24;
        $col = chr(64 + $a); // A
        foreach ($pl["invoiceform"]['b'.$b] as $item => $value) {
            if (($item > 3)&&($b == 2)) {
                $sheet->insertNewRowBefore($row, 1);
            }
            $sheet->setCellValue($col.$row, $value);
            $row++;
        }
    }
//    N行后3列
    for ($a = 1, $b = 14; $a <= 3  ; $a++, $b++) {
        $row = 24;
        $col = chr(77 + $a); // A
        foreach ($pl["invoiceform"]['b'.$b] as $item => $value) {
            $sheet->setCellValue($col.$row, $value);
            $row++;
        }
    }

}



$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

unset($_SESSION['packinglist'] ); //注销SESSION

$filenameout = "PackingList_{$pl['shortname']}_{$pl['invoiceno']}";
outExcel($spreadsheet,$filenameout);

