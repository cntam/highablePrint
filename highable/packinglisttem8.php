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
//var_dump($pl);
//$temno = $pl["temno"];
//$titlearr = unserialize(gzuncompress(base64_decode($pl["cctitle"])));
//print_r($titlearr);

//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/packinglisttem8.xlsx');
$sheet = $spreadsheet->getActiveSheet();

$spreadsheet->getActiveSheet()->setTitle("sheet1");
//$sheet->setCellValue('A1', 'Hello World !');
$spreadsheet->getDefaultStyle()->getFont()->setName('微软雅黑');
$spreadsheet->getDefaultStyle()->getFont()->setSize(12);
//$spreadsheet->getActiveSheet()->getDefaultRowDimension()->setRowHeight(50); //行默认高度
//$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(20);  //列宽度
//$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(50);  //列宽度
//$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(50);  //列宽度
//
//$spreadsheet->getActiveSheet()->getRowDimension('1')->setRowHeight(36); //列高度
//$spreadsheet->getActiveSheet()->getRowDimension('2')->setRowHeight(50); //列高度




setMergeCells($sheet,"B8:C8","B8",$pl["invoicedata"]['a1'],$noborderLeft);
setMergeCells($sheet,"B9:C9","B9",$pl["invoicedata"]['a2'],$noborderLeft);
setMergeCells($sheet,"B10:C10","B10",$pl["invoicedata"]['a3'],$noborderLeft);
setMergeCells($sheet,"B11:C11","B11",$pl["invoicedata"]['a4'],$noborderLeft);

setCell($sheet,"O6",$pl["invoicedata"]['a8'],$noborderLeft);


setMergeCells($sheet,"A17:F17","A17",$pl["invoicedata"]['a5'],$noborderLeft);
setMergeCells($sheet,"A18:C18","A18",'STYLE  NO:'.$pl["invoicedata"]['a6'],$noborderLeft);
setMergeCells($sheet,"D18:F18","D18",'STYLE CODE:'.$pl["invoicedata"]['a7'],$noborderLeft);


/**
COLOUR & SIZE BREAKDOWN FOR U.K ORDER
 *
 */
//$col = 'E';
//$row = 37;
//for($i=1,$r=0;$i<7;$i++,$r++) {
//    $avalue = $pl["invoiceform"]["b1"][$r];
//
//    setCell($sheet, $col.$row, $avalue, $noborderCenter);
//    $col++;
//}
setCell($sheet, 'K39', $pl["invoiceform"]["ba1"][6], $noborderCenter);
//第二行
//        $u = ($costp2["alist"]["a10"] - 1);$u >= 0;$u--



if ($pl["invoiceform"]["ba1"][4] > 0) {
    $row = 38;

        for ($i = 0, $v = 1; $v <= $pl["invoiceform"]["ba1"][4]; $i++, $v++) {
            $b = 18;
            $col = 'E';

            if($i >0){
                $row++;
                $spreadsheet->getActiveSheet()->insertNewRowBefore($row, 1);
            }




            for($t = 1;$t<=6;$t++){
                $avalue = $pl['invoiceform']['b'.$b][$i];
                setCell($sheet, $col . $row, $avalue, $noborderCenter);
                $b++;
                $col++;
            }
            setCell($sheet, 'B'. $row, $pl['invoiceform']['b17'][$i], $noborderCenter);
            setCell($sheet, 'K'. $row, $pl['invoiceform']['b26'][$i], $noborderCenter);
        }

}


/**
 *   C/NO.
 */
$col = 'E';
$row = 21;
for($i=1,$r=0;$i<7;$i++,$r++) {
    $avalue = $pl["invoiceform"]["b1"][$r];

    setCell($sheet, $col.$row, $avalue, $noborderCenter);
    $col++;
}

//TOTAL:
setCell($sheet, 'B32', $pl["invoiceform"]["ba1"][0], $noborderCenter);
setCell($sheet, 'K30', $pl["invoiceform"]["ba1"][1], $noborderCenter);
setCell($sheet, 'M30', $pl["invoiceform"]["ba1"][2], $noborderCenter);
setCell($sheet, 'N30', $pl["invoiceform"]["ba1"][3], $noborderCenter);

setMergeCells($sheet,"A35:F35","A35",$pl["invoiceform"]["b6"],$noborderLeft);



$row = 21;
if ($pl["invoiceform"]["brownum"] > 0) {
    for ($i = 0, $v = 1; $v <= $pl["invoiceform"]["brownum"]; $i++, $v++) {
        $col = 'A';
        $b = 2;
        $row++;
        if($v > 8){
            //$row = 30;
            $spreadsheet->getActiveSheet()->insertNewRowBefore($row, 1);
        }

        for($t = 1;$t<=15;$t++){

            $avalue = $pl['invoiceform']['b'.$b][$i];
            setCell($sheet, $col . $row, $avalue, $noborderCenter);

            if($t == 5){
               $b++;
            }elseif ($t == 12){
                $b++;
                $col++;
                $col++;
            }else{
                $b++;
                $col++;
            }
        }


    }
}




// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$spreadsheet->setActiveSheetIndex(0);

unset($_SESSION['packinglist'] ); //注销SESSION

//$spreadsheet->getActiveSheet()->getPageSetup()
//    ->setOrientation(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE);  //横放置
$spreadsheet->getActiveSheet()->getPageSetup()
    ->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);  //A4
$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

$filenameout = "PackingList_{$pl['shortname']}";
outExcel($spreadsheet,$filenameout);



