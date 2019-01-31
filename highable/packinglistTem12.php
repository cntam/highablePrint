<?php
require_once ('aidenfunc.php');
$pl =  $_SESSION['packinglist'];


use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/*
 * 思路 先填固定行 后增加 可变行
 * 1
 */


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/packinglisttem12.xlsx');
$sheet = $spreadsheet->getActiveSheet();

$spreadsheet->getActiveSheet()->setTitle("sheet1");
//$sheet->setCellValue('A1', 'Hello World !');
$spreadsheet->getDefaultStyle()->getFont()->setName('微软雅黑');
$spreadsheet->getDefaultStyle()->getFont()->setSize(12);
//$spreadsheet->getActiveSheet()->getDefaultRowDimension()->setRowHeight(50); //行默认高度
$spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(10);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(10);  //列宽度
//$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(50);  //列宽度
//
$spreadsheet->getActiveSheet()->getRowDimension('1')->setRowHeight(40); //列高度
//$spreadsheet->getActiveSheet()->getRowDimension('2')->setRowHeight(50); //列高度



setCell($sheet,"B3",$pl["invoicedata"]['invoiceNumber'],$noborderLeft);
setCell($sheet,"B6",$pl["invoicedata"]['a1'],$noborderLeft);
setMergeCells($sheet,"B4:H4","B4",$pl["invoicedata"]['a2'],$noborderLeft);
setCell($sheet,"B5",$pl["invoicedata"]['a3'],$noborderLeft);
setCell($sheet,"D5",$pl["invoicedata"]['a4'],$noborderLeft);

//TOTAL:
setCell($sheet,"E16",$pl["invoicedata"]['a5'],$noborderLeft);
setCell($sheet,"F16",$pl["invoicedata"]['a6'],$noborderLeft);
setCell($sheet,"G16",$pl["invoicedata"]['a7'],$noborderLeft);


$row = 8;
if ($pl["invoiceform"]["brownum"] > 0) {
    for ($i = 0, $v = 1; $v <= $pl["invoiceform"]["brownum"]; $i++, $v++) {
        $col = 'A';
        $b = 1;
        $row++;
        if($v > 6){
            //$row = 30;
            $spreadsheet->getActiveSheet()->insertNewRowBefore($row, 1);
        }

        for($t = 1;$t<=8;$t++){

            $avalue = $pl['invoiceform']['b'.$b][$i];
            setCell($sheet, $col . $row, $avalue, $noborderCenter);

            $b++;
            $col++;
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


$filenameout = "PackingList_{$pl['shortname']}";
outExcel($spreadsheet,$filenameout);



