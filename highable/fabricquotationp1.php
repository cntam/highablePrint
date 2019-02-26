<?php
require_once 'aidenfunc.php';

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器




$fabp1 =   $_SESSION['fabricquotationp1'];
//var_dump($fabp1);

//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/fabricquotationp1.xlsx');
$sheet = $spreadsheet->getActiveSheet();

$spreadsheet->getActiveSheet()->setTitle("第一页");

$spreadsheet->getDefaultStyle()->getFont()->setName('微软雅黑');
$spreadsheet->getDefaultStyle()->getFont()->setSize(12);
//$spreadsheet->getActiveSheet()->getDefaultRowDimension()->setRowHeight(50);


//$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(15);  //列宽度
//$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(10);  //列宽度
//$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(15);  //列宽度

$colarr = range("A","Z");
for($k=0;$k<count($fabp1['title']['a1']);$k++){
    $spreadsheet->getActiveSheet()->getColumnDimension($colarr[$k])->setWidth(15);  //列宽度
}
//$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(15);  //列宽度

$spreadsheet->getActiveSheet()->getPageMargins()->setRight(0.1);
$spreadsheet->getActiveSheet()->getPageMargins()->setLeft(0.1); //页边距


//$spreadsheet->getActiveSheet()->setCellValue('C4', $fabp1["alist"]['a1']);
$spreadsheet->getActiveSheet()->setCellValue('E1', 'DATE: '.$fabp1["date"]);
//$spreadsheet->getActiveSheet()->setCellValue('A4', $fabp1["quotitle"]);

/**
 *  title
 */
//setCell($sheet,'A4',$fabp1["quotitle"],$boldfontbordersLeft);

$col = chr(65 + count($fabp1['title']['a1']) - 1 );
setMergeCells($sheet,"A4:{$col}4",'A4',$fabp1["quotitle"],$boldfontbordersLeft);

$spreadsheet->getActiveSheet()->setCellValue('A1', $fabp1["alist"]["head"]);
$row = 6;
/**
 * 标题
 */
$a = 0;
foreach ($fabp1['title']['a1'] as $value){

    $col = chr(65 + $a);
    $colname = $col.$row;
    $spreadsheet->getActiveSheet()->setCellValue($colname, $value);
    $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($boldfontbordersLeft);
    $a++;
}

/**
 * alist
 */
$row++;
for ($y = 0, $i = 1; $i <= $fabp1["alist"]['alistnum']; $i++, $y++) {

    $tdHTML = '';

    for($u = 0,$n = 1;$u< count($fabp1['title']['a1']);$u++,$n++){
        $col = chr(65 + $u);
        if($u == 3){

            $thisvalue = $fabp1["alist"]['a'.$n][$y];
            $n++;
            $issel =  $fabp1["alist"]['a'.$n][$y] == '1' ?  "IN" :  "CM" ;
            $thisvalue .= ' '.$issel;
            $spreadsheet->getActiveSheet()->setCellValue($col.$row, $thisvalue);
            $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($bordersLeft);
        }elseif ($u == 4){
            $thisvalue = $fabp1["alist"]['a'.$n][$y];
            $n++;
            //$issel =  $fabp1["alist"]['a'.$n][$y] == '1' ?  "G/M2" :  "G/Y" ;
            $issel =  $fabp1["alist"]['a'.$n][$y];
            $thisvalue .= ' '.$issel;
            $spreadsheet->getActiveSheet()->setCellValue($col.$row, $thisvalue);
            $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($bordersLeft);
        }else{
            $spreadsheet->getActiveSheet()->setCellValue($col.$row, $fabp1["alist"]['a'.$n][$y]);
            $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($bordersLeft);
        }

    }

    $row++;
}

$row = $row>20 ? $row : 20;
/**
 * remark
 */
$spreadsheet->getActiveSheet()->mergeCells("B{$row}:E{$row}");
$spreadsheet->getActiveSheet()->setCellValue('B'.$row, $fabp1["alist"]['remarks']);
$spreadsheet->getActiveSheet()->getStyle("B{$row}:E{$row}")->applyFromArray($noborderLeft);
$row++;
if($fabp1["blist"]['blistnum'] > 0){
    foreach ($fabp1["blist"]['b1'] as $value){
        $spreadsheet->getActiveSheet()->mergeCells("B{$row}:E{$row}");
        $spreadsheet->getActiveSheet()->setCellValue('B'.$row, $value);
        $spreadsheet->getActiveSheet()->getStyle("B{$row}:E{$row}")->applyFromArray($noborderLeft);
        $row++;
    }
}



$spreadsheet->getActiveSheet()->getPageSetup()
    ->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);//打印 A4
$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页


// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$spreadsheet->setActiveSheetIndex(0);

unset($_SESSION['fabricquotationp1'] ); //注销SESSION

$filenameout = "Fabric_Quotation_{$fabp1["alist"]['head']}_";
outExcel($spreadsheet,$filenameout);

