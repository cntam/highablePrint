<?php
header("Content-type: text/html; charset=utf-8");
require_once 'aidenfunc.php';

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器


$fabp1 =   $_SESSION['stocklist'];
//var_dump($fabp1);

$spreadsheet = new Spreadsheet();
//$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/fabricquotationp1.xlsx');
$sheet = $spreadsheet->getActiveSheet();

$sheet->setTitle("sheet1");

$spreadsheet->getDefaultStyle()->getFont()->setName('微软雅黑');
$spreadsheet->getDefaultStyle()->getFont()->setSize(12);
//$sheet->getColumnDimension('F')->setWidth(18);  //列宽度 修改


$spreadsheet->getActiveSheet()->getRowDimension('1')->setRowHeight(30); //列高度

$sheet->getPageMargins()->setRight(0.1);
$sheet->getPageMargins()->setLeft(0.1); //页边距


//填数据 修改
if ($fabp1['action'] == 'material') {
    $title = 'MATERIAL STOCK REPORT';
} else {
    $title = 'FABRIC STOCK REPORT';
}

setMergeCells($sheet,'A1:G1','A1',$title,$noborderboldfontCenter);
setMergeCells($sheet,'E2:F2','E2','LAST UPDATE DATE:',$noborderboldfontLeft);

$sheet->setCellValue('A3', '客户');
$sheet->getStyle('A3')->applyFromArray($boldfontbordersLeft);

$row = 3;
/**
 * 标题
 */
$colarr = range("A","Z");

$a = 0;
foreach ($fabp1['title'] as $item=>$value){

    $sheet->getColumnDimension($colarr[$item])->setWidth(15);  //列宽度

    if($value == '(+/-) FUNCTION'){
        continue;
    }
    $col = chr(66 + $a); // B
    $colname = $col.$row;
    $sheet->setCellValue($colname, $value);
    $sheet->getStyle($col.$row)->applyFromArray($boldfontbordersLeft);
    $a++;
}

/**
 * content
 */
$a = 0;

foreach ($fabp1["info"] as $value) {
    $row = 4;
    $col = chr(65 + $a); // A
//    $colname = $col.$row;
    foreach ($value as $childvalue) {
       $contentval = stripcslashes($childvalue);

        setCell($sheet,$col.$row,$contentval,$bordersLeft);
//        $sheet->getStyle($col.$row)->applyFromArray($styleArray);
//        $sheet->setCellValue($col.$row, $childvalue);
        $row++;
    }
    $a++;
}



$sheet->getPageSetup()->setFitToPage(true); //将工作表调整为一页


// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$spreadsheet->setActiveSheetIndex(0);

unset($_SESSION['stocklist'] ); //注销SESSION



// 修改 根据不同action
if ($fabp1['action'] == 'material') {
    $filenameout = 'Material_Stock_List_';
} else if ($fabp1['action'] == 'fabric') {
    $filenameout = 'Fabric_Stock_List_';
}

outExcel($spreadsheet,$filenameout);

