<?php
header("Content-type: text/html; charset=utf-8");
require_once 'aidenfunc.php';
// modified by fa at 2019.01.16
// 友聯廠
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;

$potem2 =  $_SESSION['potem2'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/potem2.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle("sheet1");
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
//$sheet->getDefaultColumnDimension()->setWidth(20);  //设置默认列宽
$sheet->getColumnDimension('A')->setWidth(20);  //列宽度
$sheet->getColumnDimension('B')->setWidth(20);  //列宽度
$sheet->getColumnDimension('C')->setWidth(20);  //列宽度
$sheet->getColumnDimension('D')->setWidth(20);  //列宽度
$sheet->getColumnDimension('E')->setWidth(20);  //列宽度
$sheet->getColumnDimension('F')->setWidth(20);  //列宽度
$sheet->getColumnDimension('G')->setWidth(20);  //列宽度
//$sheet->getColumnDimension('H')->setWidth(15);  //列宽度
//$sheet->getColumnDimension('I')->setWidth(15);  //列宽度
$spreadsheet->getDefaultStyle()->getFont()->setSize(7);

$styleArray1 = [
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
        'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
        'wrapText' => true,
        'ShrinkToFit'=>true,
    ],
    'font' => [
        'Size' => '8',
    ],

    'borders' => [
        'top' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
        'bottom' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
        'left' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
        'right' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],

    ],
    'font' => [
        'Size' => '8',
    ],

];

//填数据
//poheader
setCell($sheet, "A1", $potem2["remark"]["poheader"]["poheada1"], $noborderCenter);
setCell($sheet, "A2", 'Address:'.$potem2["remark"]["poheader"]["poheada2"], $noborderCenter);
setCell($sheet, "A3", $potem2["remark"]["poheader"]["poheada3"], $noborderCenter);
setCell($sheet, "A4", $potem2["remark"]["poheader"]["poheada4"].' '.$potem2["remark"]["poheader"]["poheada5"], $noborderCenter);
setCell($sheet, "A5", '', $noborderCenter);
//setCell($sheet, "A5", $potem2["remark"]["poheader"]["poheada6"], $noborderCenter);

$sheet->setCellValue('A10', 'FM:');

$sheet->setCellValue('B7', $potem2["tosb"]);
$sheet->setCellValue('G7', $potem2 ["podate"]);
$sheet->setCellValue('B8', $potem2["toaddr"]["a1"]);
$sheet->setCellValue('B9', $potem2["toaddr"]["a2"].'  FAX: '.$potem2["toaddr"]["a3"]);
$sheet->setCellValue('B10', $potem2["toaddr"]["a4"]);
$sheet->setCellValue('G10', $potem2["toaddr"]["a5"]);


//中部form

$nowcol = 15;
$sheet->mergeCells("A{$nowcol}:F{$nowcol}");
$sheet->setCellValue('A'.$nowcol, '(PO NO:  '.$potem2["orderform"]["midpono"].' 注：請在開發票時把"PO NO"寫上，不可重複)');
//$sheet->setCellValue('I'.$nowcol, $potem2["invoiceform"]["amout"]);
//
$nowcol++;
$nowcol++;

for($x = 0 ,$c = 1; $x <= $potem2["orderform"]["formnum"]; $x++ ,$c++){

$f19 = 17 + 1 * $x;

$sheet->mergeCells("B{$f19}:E{$f19}");

$formarr = array('A'.$f19,'B'.$f19,'F'.$f19,'G'.$f19);

    for($i = 1,$y = 0; $i <= $potem2["orderform"]["brrnum"] ; $i++ ,$y++){

        $sheet->setCellValue($formarr[$y],  $potem2["orderform"]['b'.$i][$x]);

    }


    $nowcol = 17  +   1 * $c;

    $sheet->getStyle('A'.$f19)->applyFromArray($styleArray1);
    $sheet->getStyle("B{$f19}:E{$f19}")->applyFromArray($styleArray1);
    $sheet->getStyle('F'.$f19)->applyFromArray($styleArray1);
    $sheet->getStyle('G'.$f19)->applyFromArray($styleArray1);

    if($x >6){
        $sheet->insertNewRowBefore($nowcol, 1);
    }

}
$nowcol = $potem2["orderform"]["formnum"] > 6 ? ($nowcol + 2) : 27;
//$sheet->getCell('A1')->setValue($nowcol); 貨送以下地址
$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
$sheet->setCellValue('A'.$nowcol, '貨送以下地址');
$nowcol++;

$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
$sheet->setCellValue('A'.$nowcol, $potem2["remark"]["c1"]);
$nowcol++;


$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
$sheet->setCellValue('A'.$nowcol, $potem2["remark"]["c2"]);
$nowcol++;
$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
$sheet->setCellValue('A'.$nowcol, $potem2["remark"]["c3"]);

$nowcol++;
$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
$sheet->setCellValue('A'.$nowcol, 'REAMRK:'.$potem2["remark"]["c4"]);
//setCell($sheet,'A'.$nowcol, 'REAMRK:'.$potem2["remark"]["c4"], $noborderLeft);

$sheet->getPageSetup()
    ->setOrientation(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_PORTRAIT);  //竖放置
$sheet->getPageSetup()
    ->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);  //A4
$sheet->getPageSetup()->setFitToPage(true); //将工作表调整为一页

unset($_SESSION['potem2'] ); //注销SESSION

$filenameout = 'PO_'.$potem2['pono'];
outExcel($spreadsheet,$filenameout);
