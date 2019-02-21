<?php
require_once 'aidenfunc.php';
header("Content-type: text/html; charset=utf-8");


$potem15 =  $_SESSION['potem15'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/potem15.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle("sheet1");
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
//$sheet->getDefaultColumnDimension()->setWidth(30);  //设置默认列宽
//for($j=0;$j<=6;$j++){
//    $col = chr(65 + $j);
//    $sheet->getColumnDimension($col)->setWidth(20);  //列宽度
//}

$sheet->getColumnDimension('A')->setWidth(25);  //列宽度
$sheet->getColumnDimension('B')->setWidth(15);  //列宽度
$sheet->getColumnDimension('C')->setWidth(20);  //列宽度
$sheet->getColumnDimension('D')->setWidth(20);  //列宽度
$sheet->getColumnDimension('E')->setWidth(10);  //列宽度
$sheet->getColumnDimension('F')->setWidth(15);  //列宽度
$sheet->getColumnDimension('G')->setWidth(20);  //列宽度
$sheet->getColumnDimension('H')->setWidth(20);  //列宽度
$sheet->getColumnDimension('I')->setWidth(30);  //列宽度

$spreadsheet->getDefaultStyle()->getFont()->setSize(7);

//填数据
$sheet->setCellValue('B6', $potem15["tosb"]);
$sheet->setCellValue('B8', $potem15["podate"]);
////
//$sheet->setCellValue('G8', $potem15["toaddr"]["a1"]);
//$sheet->setCellValue('B9', $potem15["toaddr"]["a2"]);
//$sheet->setCellValue('B11', $potem15["toaddr"]["a3"]);
//$sheet->setCellValue('B12', $potem15["toaddr"]["a4"]);
$toaddr = array('H6','B7','H7','H8','B9','H9','B12','B13','B14','B15','B16','B17','B18','B19','B20','G12','G13','G14','G15','G16','G17','G18','G19','G20');  //

for($i = 1,$y = 0; $i <= count($toaddr) ; $i++ ,$y++){

    $sheet->setCellValue($toaddr[$y],  $potem15["toaddr"]["a".$i]);

}
//
//
////中部form

//$nowcol = 24;
//////$sheet->mergeCells("A{$nowcol}:F{$nowcol}");
//$sheet->setCellValue('B'.$nowcol, $potem15["orderform"]["midpono"]);
//////$sheet->setCellValue('I'.$nowcol, $potem15["invoiceform"]["amout"]);
//
//
$sheet->setCellValue('D23', "Unit Price". "(".$potem15["orderform"]["b5"].")");
for($x = 0 ,$c = 1; $c <= $potem15["orderform"]["formnum"]; $x++ ,$c++){

$f19 = 24 + 1 * $x;

//    $sheet->mergeCells("A{$f19}:B{$f19}");
//    $sheet->mergeCells("C{$f19}:G{$f19}");


$formarr = array('A'.$f19,'B'.$f19,'C'.$f19,'D'.$f19);

    for($i = 1,$y = 0; $i <= 4 ; $i++ ,$y++){

        $sheet->setCellValue($formarr[$y],  $potem15["orderform"]['b'.$i][$x]);

    }

    $nowcol = 24  +  1 * $c;


    if($x >15){
        $sheet->insertNewRowBefore($nowcol, 1);
    }

}
//$nowcol = $potem15["orderform"]["formnum"] > 4 ? ($nowcol + 1) : 21;
//$sheet->setCellValue('C'.$nowcol, $potem15["toaddr"]["a5"]);
//$nowcol++;
//
//$sheet->setCellValue('C'.$nowcol, $potem15["toaddr"]["a6"]);
//$nowcol++;
//$sheet->setCellValue('C'.$nowcol, $potem15["toaddr"]["a7"]);
//$nowcol++;
//$sheet->setCellValue('A'.$nowcol, $potem15["toaddr"]["a8"]);
//$nowcol++;
//$nowcol++;
//$nowcol++;
//$nowcol++;
//$nowcol++;
//$nowcol++;
//$nowcol++;
//
//$sheet->setCellValue('C'.$nowcol, $potem15["remark"]["c1"]);
//$nowcol++;
//$sheet->setCellValue('C'.$nowcol, $potem15["remark"]["c2"]);

////
////$sheet->setCellValue('B'.$nowcol, 'E-MAIL (電郵): '.$potem15["remark"]["c3"]);
////$sheet->setCellValue('I'.$nowcol, $potem15["remark"]["c4"]);
//////$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
//$sheet->setCellValue('O'.$nowcol, $potem15["remark"]["c3"]);
//$nowcol++;
//
//$sheet->setCellValue('O'.$nowcol, $potem15["remark"]["c4"]);

$sheet->getPageSetup()
    ->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);  //A4
$sheet->getPageSetup()->setFitToPage(true); //将工作表调整为一页

unset($_SESSION['potem15'] ); //注销SESSION

$filenameout = 'PO_'.$potem15['pono'];
outExcel($spreadsheet,$filenameout);

