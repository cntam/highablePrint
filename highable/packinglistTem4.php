<?php
session_start();
header("Content-type: text/html; charset=utf-8");
//Modified by 俊伟  /*港源行國際有限公司*/


require_once('autoloadconfig.php');  //判断是否在线

require_once ('img.php');

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;

$packinglistTem4 =  $_SESSION['packinglist'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/packinglistTem4.xlsx');
$spreadsheet->getActiveSheet()->setTitle("sheet1");
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
$sheet->setCellValue('A3', 'Invoice No.:'.$packinglistTem4["invoicedata"]["invoiceNumber"]);
$sheet->setCellValue('H3', 'Date: '.$packinglistTem4["invoicedata"]["a40"]);
$sheet->setCellValue('T3', 'Production Order No.: '.$packinglistTem4["invoicedata"]["a1"]);

$sheet->setCellValue('A6', $packinglistTem4["invoicedata"]["a2"]);
$sheet->setCellValue('A8', $packinglistTem4["invoicedata"]["a5"]);

$sheet->setCellValue('R6', $packinglistTem4["invoicedata"]["a3"]);
$sheet->setCellValue('A7', $packinglistTem4["invoicedata"]["a4"]);
$sheet->setCellValue('A9', $packinglistTem4["invoicedata"]["a6"]);

$sheet->setCellValue('R13', $packinglistTem4["invoicedata"]["a7"]);
$sheet->setCellValue('R14', $packinglistTem4["invoicedata"]["a8"]);
$sheet->setCellValue('R15', $packinglistTem4["invoicedata"]["a9"]);
$sheet->setCellValue('R16', $packinglistTem4["invoicedata"]["a10"]);

$sheet->setCellValue('D20', $packinglistTem4["invoicedata"]["a11"]);
$sheet->setCellValue('D21', $packinglistTem4["invoicedata"]["a12"]);

//footer
$sheet->setCellValue('A65', $packinglistTem4["invoicedata"]["a39"]);

$sheet->setCellValue('G62', $packinglistTem4["invoicedata"]["a36"]);
$sheet->setCellValue('K62', $packinglistTem4["invoicedata"]["a37"]);
$sheet->setCellValue('O62', 'H      '.$packinglistTem4["invoicedata"]["a38"]);

//form
//size格
$sheet->setCellValue('U23', $packinglistTem4["invoicedata"]["a13"]);
$sheet->setCellValue('V23', $packinglistTem4["invoicedata"]["a14"]);
$sheet->setCellValue('W23', $packinglistTem4["invoicedata"]["a15"]);
$sheet->setCellValue('X23', $packinglistTem4["invoicedata"]["a16"]);
$sheet->setCellValue('Y23', $packinglistTem4["invoicedata"]["a17"]);
$sheet->setCellValue('Z23', $packinglistTem4["invoicedata"]["a18"]);
$sheet->setCellValue('AA23', $packinglistTem4["invoicedata"]["a19"]);
$sheet->setCellValue('AB23', $packinglistTem4["invoicedata"]["a20"]);
$sheet->setCellValue('AC23', $packinglistTem4["invoicedata"]["a21"]);

//form 动态1
$row = 34;
foreach ($packinglistTem4["invoiceform"]["b18"] as $item => $value) {
    if ($item > 4) {
        $sheet->insertNewRowBefore($row, 5);
//        固定文字
        $sheet->setCellValue('A'.($row + 1), 'TOTAL GROSS Wt:');
        $sheet->setCellValue('E'.($row + 1), 'Kg.');
        $sheet->setCellValue('A'.($row + 3), 'TOTAL NET Wt:');
        $sheet->setCellValue('E'.($row + 3), 'Kg.');
    }
    $sheet->setCellValue('A'.$row, $value);
    $row += 5;
}
$row = 35;
foreach ($packinglistTem4["invoiceform"]["b19"] as $item => $value) {
    $sheet->setCellValue('G'.$row, $value);
    $contextG = 'G'.$row;
    $contextH = 'H'.$row;
    $sheet->mergeCells("$contextG:$contextH");
    $sheet->getStyle('G'.$row)->applyFromArray($styleArray1);
    $sheet->getStyle('H'.$row)->applyFromArray($styleArray1);
    $row += 5;
}
$row = 37;
foreach ($packinglistTem4["invoiceform"]["b20"] as $item => $value) {
    $sheet->setCellValue('G'.$row, $value);
    $contextG = 'G'.$row;
    $contextH = 'H'.$row;
    $sheet->mergeCells("$contextG:$contextH");
    $sheet->getStyle('G'.$row)->applyFromArray($styleArray1);
    $sheet->getStyle('H'.$row)->applyFromArray($styleArray1);
    $row += 5;
}

//total3行
if ($packinglistTem4["invoiceform"]["brownum"] > 0) {
    $arr = array('U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD');
    $row = 31;
    for ($a = 0, $b = 23; $a < count($arr) ; $a++, $b++) {
        $sheet->setCellValue($arr[$a].$row, $packinglistTem4["invoicedata"]['a'.$b]);
    }
    $sheet->insertNewRowBefore($row, 1);
    $contextArr = array('A'.$row,'B'.$row, 'C'.$row, 'E'.$row, 'F'.$row, 'H'.$row, 'I'.$row, 'O'.$row, 'P'.$row,
        'Q'.$row, 'R'.$row, 'S'.$row, 'AD'.$row, 'AE'.$row);
    $sheet->mergeCells("$contextArr[0]:$contextArr[1]");
    $sheet->mergeCells("$contextArr[2]:$contextArr[3]");
    $sheet->mergeCells("$contextArr[4]:$contextArr[5]");
    $sheet->mergeCells("$contextArr[6]:$contextArr[7]");
    $sheet->mergeCells("$contextArr[8]:$contextArr[9]");
    $sheet->mergeCells("$contextArr[10]:$contextArr[11]");
    $sheet->mergeCells("$contextArr[12]:$contextArr[13]");
    $row = 30;
    $sheet->setCellValue('I'.$row, $packinglistTem4["invoiceform"]["b21"][0]);
    for ($a = 0, $b = 22; $a < count($arr) ; $a++, $b++) {
        $sheet->setCellValue($arr[$a].$row, $packinglistTem4["invoiceform"]['b'.$b][0]);

    }
    $row = 31;
    $sheet->setCellValue('I'.$row, $packinglistTem4["invoiceform"]["b21"][1]);
    for ($a = 0, $b = 22; $a < count($arr) ; $a++, $b++) {
        $sheet->setCellValue($arr[$a].$row, $packinglistTem4["invoiceform"]['b'.$b][1]);
    }

}

if ($packinglistTem4["invoiceform"]["brownum"] > 0) {
    $arr = array('A', 'C', 'F', 'I', 'P', 'R', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD');

    for ($a = 0, $b = 1; $a < count($arr); $a++, $b++) {
        $row = 24;
        foreach ($packinglistTem4["invoiceform"]['b'.$b] as $item=>$value) {
            if (($item > 4)&&($b == 1)) {
                $sheet->insertNewRowBefore($row, 1);
                $contextArr = array('A'.$row,'B'.$row, 'C'.$row, 'E'.$row, 'F'.$row, 'H'.$row, 'I'.$row, 'O'.$row, 'P'.$row,
                    'Q'.$row, 'R'.$row, 'S'.$row, 'AD'.$row, 'AE'.$row);
                $sheet->mergeCells("$contextArr[0]:$contextArr[1]");
                $sheet->mergeCells("$contextArr[2]:$contextArr[3]");
                $sheet->mergeCells("$contextArr[4]:$contextArr[5]");
                $sheet->mergeCells("$contextArr[6]:$contextArr[7]");
                $sheet->mergeCells("$contextArr[8]:$contextArr[9]");
                $sheet->mergeCells("$contextArr[10]:$contextArr[11]");
                $sheet->mergeCells("$contextArr[12]:$contextArr[13]");
            }
            $sheet->setCellValue($arr[$a].$row, $value);
            $row++;
        }
    }

}


$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

//unset($_SESSION['potem1'] ); //注销SESSION

$output=  ($_GET['action'] == 'formdown' )? 1:0;
$nt = date("md",time()); //转换为日期。
$filenameout = 'Packinglist_MCQ_'.$nt.'.xlsx';
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

