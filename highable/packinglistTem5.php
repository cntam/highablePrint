<?php
//Modified by 俊伟
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
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/packinglistTem5.xlsx');
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
//$spreadsheet->getActiveSheet()->getDefaultRowDimension()->setRowHeight(50); //行默认高度
//$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(20);  //列宽度
//$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(50);  //列宽度
//$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(50);  //列宽度
//
//$spreadsheet->getActiveSheet()->getRowDimension('1')->setRowHeight(36); //列高度
//$spreadsheet->getActiveSheet()->getRowDimension('2')->setRowHeight(50); //列高度

require_once ('aidenfunc.php');



setMergeCells($sheet,"C7:D7","C7",$pl["invoicedata"]['a2'],$noborderLeft);
setMergeCells($sheet,"C8:E8","C8",$pl["invoicedata"]['a3'],$noborderLeft);
setMergeCells($sheet,"C9:E9","C9",$pl["invoicedata"]['a4'],$noborderLeft);
setMergeCells($sheet,"C10:E10","C10",$pl["invoicedata"]['a5'],$noborderLeft);
//
//setMergeCells($sheet,"W13:X13","W13",$pl["invoicedata"]['a1'],$noborderLeft);
setCell($sheet,"W13", $pl["invoicedata"]['a1'],$noborderLeft);
setMergeCells($sheet,"B14:E14","B14",$pl["invoicedata"]['a6'],$noborderLeft);

//TOTAL
setMergeCells($sheet,"B40:J40","B40",$pl["invoicedata"]['a23'],$noborderLeft);

//SUMMARY OF TOTAL BREAKDOWN表格固定
$col = 'G';
$row = 30;
for($i=1,$r=8;$i<=7;$i++,$r++) {
    $avalue = $pl["invoicedata"]['a'.$r];

    setCell($sheet, $col.$row, $avalue, $noborderCenter);
    $col++;
}

setCell($sheet,"S38",$pl["invoicedata"]['a21'],$noborderLeft);
setCell($sheet,"T38",'PCS',$noborderLeft);

//C/NO.表格固定
$col = 'G';
$row = 16;
for($i=1,$r=7;$i<=7;$i++,$r++) {
    $avalue = $pl["invoicedata"]['a'.$r];

    setCell($sheet, $col.$row, $avalue, $noborderCenter);
    $col++;
}

setCell($sheet,"Q28",$pl["invoicedata"]['a25'],$noborderLeft);
setCell($sheet,"S28",$pl["invoicedata"]['a26'],$noborderLeft);
setCell($sheet,"U28",$pl["invoicedata"]['a27'],$noborderLeft);
setCell($sheet,"T28",'"',$noborderLeft);
setCell($sheet,"V28",$pl["invoicedata"]['a28'],$noborderLeft);

//表格动态

//SUMMARY OF TOTAL BREAKDOWN
if ($pl["invoiceform"]["brownum"] > 0) {
    $arr = array('C', 'D', 'E', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'S', 'T');

    for ($a = 0, $b = 19; $a < count($arr); $a++, $b++) {
        $row = 31;
        foreach ($pl["invoiceform"]['b'.$b] as $item=>$value) {
            if (($item > 6)&&($b == 19)) {
                $sheet->insertNewRowBefore($row, 1);
            }
//            $sheet->setCellValue($arr[$a].$row, $value);
            setCell($sheet,$arr[$a].$row, $value, $noborderCenter);
            setCell($sheet,"R".$row, '=',$noborderCenter);
//            $sheet->getStyle('R'.$row)->applyFromArray($styleArray1);
            $row++;
        }
    }

}

//C/NO
if ($pl["invoiceform"]["brownum"] > 0) {
    $arr = array('B', 'C', 'D', 'E', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'O', 'Q', 'S', 'T', 'U', 'V', 'W');

    for ($a = 0, $b = 1; $a < count($arr); $a++, $b++) {
        $row = 17;
        foreach ($pl["invoiceform"]['b'.$b] as $item=>$value) {
            if (($item > 8)&&($b == 1)) {
                $sheet->insertNewRowBefore($row, 1);
            }
//            $sheet->setCellValue($arr[$a].$row, $value);
            setCell($sheet,$arr[$a].$row, $value, $noborderCenter);
            setCell($sheet,"N".$row, '=',$noborderCenter);
            setCell($sheet,"P".$row, 'X',$noborderCenter);
            setCell($sheet,"R".$row, '=',$noborderCenter);
            $contextArrE = 'W'.$row;
            $contextArrH = 'X'.$row;
            $sheet->mergeCells("$contextArrE:$contextArrH");
            $row++;
        }
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

$nt = date("md",time()); //转换为日期。
$filenameout = 'packlinglist_PS_'.$nt.'.xlsx';

outExcel($spreadsheet,$filenameout);



