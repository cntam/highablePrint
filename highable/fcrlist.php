<?php
require_once 'aidenfunc.php';

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器


$fabp1 =   $_SESSION['fcrlist'];
//var_dump($fabp1);

$spreadsheet = new Spreadsheet();
//$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/fabricquotationp1.xlsx');
$sheet = $spreadsheet->getActiveSheet();

$spreadsheet->getActiveSheet()->setTitle("sheet");

$spreadsheet->getDefaultStyle()->getFont()->setName('微软雅黑');
$spreadsheet->getDefaultStyle()->getFont()->setSize(12);
//$spreadsheet->getActiveSheet()->getDefaultRowDimension()->setRowHeight(50);


//$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(15);  //列宽度
//$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(10);  //列宽度
//$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(15);  //列宽度

$colarr = range("A","Z");
for($k=0;$k<count($fabp1['title']);$k++){
    $spreadsheet->getActiveSheet()->getColumnDimension($colarr[$k])->setWidth(15);  //列宽度
}
//$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(15);  //列宽度

$spreadsheet->getActiveSheet()->getPageMargins()->setRight(0.1);
$spreadsheet->getActiveSheet()->getPageMargins()->setLeft(0.1); //页边距


//$spreadsheet->getActiveSheet()->setCellValue('C4', $fabp1["alist"]['a1']);
$spreadsheet->getActiveSheet()->setCellValue('G1', 'BRAND:');
//$spreadsheet->getActiveSheet()->setCellValue('A4', $fabp1["quotitle"]);
$row = 3;
/**
 * 标题
 */
$a = 0;
foreach ($fabp1['title'] as $value){

    $col = chr(65 + $a);
    $colname = $col.$row;
    $spreadsheet->getActiveSheet()->setCellValue($colname, $value);
    $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($bordersLeft);
    $a++;
}

/**
 * content
 */
$row = 4;
for ($y = 0, $i = 1; $i <= count($fabp1["content"]); $i++, $y++) {

    $tdHTML = '';

    for($u = 0,$prt = 0,$n = 1;$n<= count($fabp1['title']);$u++,$prt++,$n++){
        $col = chr(65 + $prt);

        if($u == 1){
            setCell($sheet,$col.$row,$fabp1["content"][$y][$u],$bordersLeft);
//            $thisvalue = $fabp1["shortname"].'-'.$fabp1["content"][$y][$u];
//            $spreadsheet->getActiveSheet()->setCellValue($col.$row, $thisvalue);
//            $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($bordersLeft);
        }elseif($u == 5){

            $thisvalue = $fabp1["content"][$y][$u].' ';
            $u++;
            $thisvalue .= $fabp1["content"][$y][$u];

            setCell($sheet,$col.$row,$thisvalue,$bordersLeft);
//            $spreadsheet->getActiveSheet()->setCellValue($col.$row, $thisvalue);
//            $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($bordersLeft);
        }else{
            $spreadsheet->getActiveSheet()->setCellValue($col.$row, stripcslashes($fabp1["content"][$y][$u]));
            $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($bordersLeft);
        }

    }

    $row++;
}



$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页


// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$spreadsheet->setActiveSheetIndex(0);

unset($_SESSION['fcrlist'] ); //注销SESSION

$filenameout = "Fcrlist_";
outExcel($spreadsheet,$filenameout);
