<?php
require_once('autoloadconfig.php');  //判断是否在线
$pl =  $_SESSION['packinglist'];


use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/*
 * 思路 先填固定行 后增加 可变行
 * 1
 */
//var_dump($pl);
//$temno = $pl["temno"];
//$titlearr = unserialize(gzuncompress(base64_decode($pl["cctitle"])));
//print_r($titlearr);

//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/packinglistTem7.xlsx');
$sheet = $spreadsheet->getActiveSheet();

$spreadsheet->getActiveSheet()->setTitle("sheet1");
//$sheet->setCellValue('A1', 'Hello World !');
$spreadsheet->getDefaultStyle()->getFont()->setName('微软雅黑');
$spreadsheet->getDefaultStyle()->getFont()->setSize(12);
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

require_once ('aidenfunc.php');

// 填数据
// header
setCell($sheet,"B3", $pl["invoicedata"]['invoiceNumber'], $noborderLeft);
setCell($sheet,"B4", $pl["invoicedata"]['a1'], $noborderLeft);
setCell($sheet,"B5", $pl["invoicedata"]['a2'], $noborderLeft);
setCell($sheet,"D5", $pl["invoicedata"]['a3'], $noborderLeft);
setCell($sheet,"B6", $pl["invoicedata"]['a4'], $noborderLeft);
//footer
setCell($sheet,"E16", $pl["invoiceform"]['ba1']['0'],$noborderCenter);
setCell($sheet,"F16", $pl["invoiceform"]['ba1']['1'],$noborderCenter);
setCell($sheet,"G16", $pl["invoiceform"]['ba1']['2'],$noborderCenter);

//表格动态

if ($pl["invoiceform"]["brownum"] > 0) {
    $col = 'A';
    for ($a = 0, $b = 1; $a < 8; $a++, $b++) {
        $row = 9;
        foreach ($pl["invoiceform"]['b'.$b] as $item=>$value) {
            if (($item > 5)&&($b == 1)) {
                $sheet->insertNewRowBefore($row, 1);
            }
//            $sheet->setCellValue($arr[$a].$row, $value);
            setCell($sheet,$col.$row, $value, $noborderCenter);
            $row++;
        }
        $col++;
    }

}

// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$spreadsheet->setActiveSheetIndex(0);

//unset($_SESSION['packinglist'] ); //注销SESSION

//$spreadsheet->getActiveSheet()->getPageSetup()
//    ->setOrientation(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE);  //横放置
$spreadsheet->getActiveSheet()->getPageSetup()
    ->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);  //A4
$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页


$filenameout = 'Packlinglist_GIVENCHY_';

outExcel($spreadsheet,$filenameout);



