<?php
require_once 'aidenfunc.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;

$potem5 =  $_SESSION['potem5'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/potem5.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle("sheet1");
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
//$sheet->getDefaultColumnDimension()->setWidth(20);  //设置默认列宽
$sheet->getColumnDimension('A')->setWidth(25);  //列宽度
$sheet->getColumnDimension('B')->setWidth(25);  //列宽度
$sheet->getColumnDimension('C')->setWidth(25);  //列宽度
$sheet->getColumnDimension('D')->setWidth(30);  //列宽度
$sheet->getColumnDimension('E')->setWidth(25);  //列宽度
$sheet->getColumnDimension('F')->setWidth(25);  //列宽度
//$sheet->getColumnDimension('G')->setWidth(16);  //列宽度
//$sheet->getColumnDimension('H')->setWidth(15);  //列宽度
//$sheet->getColumnDimension('I')->setWidth(15);  //列宽度
$spreadsheet->getDefaultStyle()->getFont()->setSize(7);

//foreach (range('A','E') as $item){
//    $spreadsheet->getActiveSheet()->getColumnDimension($item)->setAutoSize(true);  //自动列宽度
//}

//填数据
//header
setCell($sheet, "A1", $potem5["remark"]["poheader"]["poheada1"], $noborderCenter);
setCell($sheet, "A2", $potem5["remark"]["poheader"]["poheada2"].' '.$potem5["remark"]["poheader"]["poheada3"], $Size12noborderCenter);
//setCell($sheet, "A4", $potem6["remark"]["poheader"]["poheada3"], $noborderCenter);
setCell($sheet, "A3", $potem5["remark"]["poheader"]["poheada4"], $noborderCenter);
//setCell($sheet, "A6", $potem6["remark"]["poheader"]["poheada6"], $noborderCenter);

$sheet->setCellValue('A6', 'TO:'.$potem5["tosb"]);
$sheet->setCellValue('E9', $potem5 ["podate"]);
$sheet->setCellValue('A7', $potem5["toaddr"]["a1"]);
$sheet->setCellValue('A8', 'TEL:'.$potem5["toaddr"]["a2"].'  FAX:'.$potem5["toaddr"]["a3"]);

$sheet->setCellValue('A9', 'E-mail:'.$potem5["toaddr"]["a4"]);
$sheet->setCellValue('A10', 'ATTN:'.$potem5["toaddr"]["a5"]);
$sheet->setCellValue('E10', $potem5["toaddr"]["a6"]);


//中部form

$nowcol = 14;
//$sheet->mergeCells("A{$nowcol}:F{$nowcol}");
$sheet->setCellValue('A'.$nowcol, '(PO NO:  '.$potem5["orderform"]["midpono"].' (注：請在送貨單和發票上注明PO NO.和OUR REF NO,不可重復,謝!)');
//$sheet->setCellValue('I'.$nowcol, $potem5["invoiceform"]["amout"]);
//
//$nowcol++;
//$nowcol++;
//
for($x = 0 ,$c = 1; $x <= $potem5["orderform"]["formnum"]; $x++ ,$c++){

$f19 = 17 + 1 * $x;

    $sheet->mergeCells("B{$f19}:C{$f19}");


$formarr = array('A'.$f19,'B'.$f19,'D'.$f19,'E'.$f19);

    for($i = 1,$y = 0; $i <= $potem5["orderform"]["brrnum"] ; $i++ ,$y++){

        $sheet->setCellValue($formarr[$y],  $potem5["orderform"]['b'.$i][$x]);

    }


    $nowcol = 17  +   1 * $c;



    if($x >4){

        $sheet->insertNewRowBefore($nowcol, 1);

    }

}
$nowcol = $potem5["orderform"]["formnum"] > 4 ? ($nowcol + 2) : 24;
//$sheet->getCell('A1')->setValue($nowcol);
////$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
////$sheet->setCellValue('A'.$nowcol, '貨送以下地址');
////$nowcol++;
//
////$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
$sheet->setCellValue('B'.$nowcol, $potem5["remark"]["c1"]);
//$nowcol++;
//$nowcol++;
//$nowcol++;
//$nowcol++;
//
////$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
//$sheet->setCellValue('D'.$nowcol, $potem5["remark"]["c2"]);
//$nowcol++;
//
////$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
//$sheet->setCellValue('A'.$nowcol, $potem5["remark"]["c3"]);
//


$sheet->getPageSetup()
    ->setOrientation(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_PORTRAIT);  //竖放置
$sheet->getPageSetup()
    ->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);  //A4
$sheet->getPageSetup()->setFitToPage(true); //将工作表调整为一页

unset($_SESSION['potem5'] ); //注销SESSION

$filenameout = 'PO_' . $potem5['shortName'].'_'.$potem5['pono'];
outExcel($spreadsheet, $filenameout);