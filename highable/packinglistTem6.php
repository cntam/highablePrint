<?php
session_start();

header("Content-type: text/html; charset=utf-8");
//Modified by 俊伟
/*港源行國際有限公司*/

require_once('autoloadconfig.php');  //判断是否在线
require_once ('img.php');

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;

$packinglistTem6 =  $_SESSION['packinglist'];

//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/packinglistTem6.xlsx');
$sheet = $spreadsheet->getActiveSheet();
//样式，下框细边
$styleArray1 = [
        'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
        'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
        'wrapText' => true,
        'ShrinkToFit'=>true,
    ],
    'borders' => [
        'bottom' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
    ],
];
//填数据
//header
//$sheet->setCellValue('P8', $packinglistTem6["invoicedata"]["invoiceNumber"]);

$sheet->setCellValue('B10', $packinglistTem6["invoicedata"]["a1"]);
$sheet->setCellValue('B11', $packinglistTem6["invoicedata"]["a2"]);
$sheet->setCellValue('B12', $packinglistTem6["invoicedata"]["a3"]);
$sheet->setCellValue('B13', 'Attn:'.$packinglistTem6["invoicedata"]["a4"]);

$sheet->setCellValue('A19', $packinglistTem6["invoicedata"]["a5"]);
$sheet->setCellValue('A20', $packinglistTem6["invoicedata"]["a6"]);
//Size Breakdown
$sheet->setCellValue('D23', $packinglistTem6["invoiceform"]["b1"]["0"]);
$sheet->setCellValue('E23', $packinglistTem6["invoiceform"]["b1"]["1"]);
$sheet->setCellValue('F23', $packinglistTem6["invoiceform"]["b1"]["2"]);
$sheet->setCellValue('G23', $packinglistTem6["invoiceform"]["b1"]["3"]);
$sheet->setCellValue('H23', $packinglistTem6["invoiceform"]["b1"]["4"]);
$sheet->setCellValue('I23', $packinglistTem6["invoiceform"]["b1"]["5"]);
$sheet->setCellValue('J23', $packinglistTem6["invoiceform"]["b1"]["6"]);
$sheet->setCellValue('K23', $packinglistTem6["invoiceform"]["b1"]["7"]);
//COLOUR/SIZE
$sheet->setCellValue('D36', $packinglistTem6["invoiceform"]["b1"]["8"]);
$sheet->setCellValue('E36', $packinglistTem6["invoiceform"]["b1"]["9"]);
$sheet->setCellValue('F36', $packinglistTem6["invoiceform"]["b1"]["10"]);
$sheet->setCellValue('G36', $packinglistTem6["invoiceform"]["b1"]["11"]);
$sheet->setCellValue('H36', $packinglistTem6["invoiceform"]["b1"]["12"]);
$sheet->setCellValue('I36', $packinglistTem6["invoiceform"]["b1"]["13"]);
$sheet->setCellValue('J36', $packinglistTem6["invoiceform"]["b1"]["14"]);
$sheet->setCellValue('K36', $packinglistTem6["invoiceform"]["b1"]["15"]);


//form total
$sheet->setCellValue('L29', $packinglistTem6["invoiceform"]["ba1"][1]);
$sheet->setCellValue('N29', $packinglistTem6["invoiceform"]["ba1"][2]);
$sheet->setCellValue('O29', $packinglistTem6["invoiceform"]["ba1"][3]);

$sheet->setCellValue('B31', $packinglistTem6["invoiceform"]["brownum"]);

//footer
//$sheet->setCellValue('C31', $packinglistTem6["invoicedata"]["invoiceNumber"]);
$sheet->setCellValue('C42', $packinglistTem6["invoiceform"]["ba1"][2]);
$sheet->setCellValue('C43', $packinglistTem6["invoiceform"]["ba1"][3]);

//form动态
if ($packinglistTem6["invoiceform"]["brownum"] > 0) {
//    COLOUR/SIZE
    for ($a = 1, $b = 18; $a <= 9 ; $a++, $b++) {
        $row = 37;
        $col = chr(67 + $a); // D
        foreach ($packinglistTem6["invoiceform"]['b'.$b] as $item => $value) {
            if (($item > 0)&&($b == 18)) {
                $sheet->insertNewRowBefore($row, 1);
            }
            $sheet->setCellValue($col.$row, $value);
            $row++;
        }
    }
    for ($a = 0; $a < $packinglistTem6["invoiceform"]["brownum"] ; $a++) {
        $row = 37;
        foreach ($packinglistTem6["invoiceform"]["b17"] as $item => $value) {
            $sheet->setCellValue('B'.$row, $value);
            $row++;
        }
    }
//    C/NO
    for ($a = 1, $b = 2; $a <= 12  ; $a++, $b++) {
        $row = 24;
        $col = chr(64 + $a); // A
        foreach ($packinglistTem6["invoiceform"]['b'.$b] as $item => $value) {
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
        foreach ($packinglistTem6["invoiceform"]['b'.$b] as $item => $value) {
            $sheet->setCellValue($col.$row, $value);
            $row++;
        }
    }

}



$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

//unset($_SESSION['potem1'] ); //注销SESSION

$output=  ($_GET['action'] == 'formdown' )? 1:0;
$nt = date("md",time()); //转换为日期。
$filenameout = 'Packinglist_GB_'.$nt.'.xlsx';
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

