<?php

require_once 'aidenfunc.php';

$potem = $_SESSION['potem13'];

//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/potem13.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle("sheet1");
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
//$sheet->getDefaultColumnDimension()->setWidth(30);  //设置默认列宽
for($j=1;$j<=25;$j++){
    $col = chr(65 + $j);
    $sheet->getColumnDimension($col)->setWidth(4);  //列宽度
}
for($j=0;$j<=10;$j++){
    $col = chr(65 + $j);
    $sheet->getColumnDimension('A'.$col)->setWidth(4);  //列宽度
}
$spreadsheet->getDefaultStyle()->getFont()->setSize(7);

//填数据
$sheet->setCellValue('H9', 'TO: '.$potem["tosb"]);

$toaddr = array('Z9','B11','T11','B13','T13','F15','O15','Y15','B23','B34','U23','B41','K41','U41','Z41','AE41','J43','AC43');  //,'C12','D12','E12','F12','G12','H12','I12','J13','K12','K13','J5'

for($i = 1,$y = 0; $i <= count($toaddr) ; $i++ ,$y++){

    $sheet->setCellValue($toaddr[$y],  $potem["toaddr"]["a".$i]);

}


//中部form

//$nowcol = 35;
////$sheet->mergeCells("A{$nowcol}:F{$nowcol}");
//$sheet->setCellValue('A'.$nowcol, 'PO NO: '.$potem["orderform"]["midpono"].'   注：請在開發票時把“PONO”寫上，不可重復，并且寫上制單號）');
////$sheet->setCellValue('I'.$nowcol, $potem["invoiceform"]["amout"]);


for($x = 0 ,$c = 1; $c <= $potem["orderform"]["formnum"]; $x++ ,$c++){

    $f19 = 44 + 1 * $x;




    $formarr = array('B'.$f19,'E'.$f19,'I'.$f19,'N'.$f19,'X'.$f19,'AC'.$f19,);

    for($i = 1,$y = 0; $i <= $potem["orderform"]["brrnum"] ; $i++ ,$y++){

        $sheet->setCellValue($formarr[$y],  $potem["orderform"]['b'.$i][$x]);

    }

    $nowcol = 44  +   1 * $c;


    if($x >4){
        $sheet->insertNewRowBefore($nowcol, 1);
    }
    $sheet->mergeCells("B{$nowcol}:D{$nowcol}");
    $sheet->mergeCells("E{$nowcol}:H{$nowcol}");
    $sheet->mergeCells("I{$nowcol}:M{$nowcol}");
    $sheet->mergeCells("N{$nowcol}:W{$nowcol}");
    $sheet->mergeCells("X{$nowcol}:AB{$nowcol}");
    $sheet->mergeCells("AC{$nowcol}:AJ{$nowcol}");
}
$nowcol = $potem["orderform"]["formnum"] > 4 ? ($nowcol + 1) : 50;
$sheet->setCellValue('Q'.$nowcol, $potem["toaddr"]["a19"]);
//////$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
$sheet->setCellValue('J18', $potem["toaddr"]["a20"]);
$nowcol++;
$nowcol++;


$sheet->setCellValue('O'.$nowcol, $potem["remark"]["c1"]);
$nowcol++;
$sheet->setCellValue('O'.$nowcol, $potem["remark"]["c2"]);
$nowcol++;
//$nowcol++;
//
//$sheet->setCellValue('B'.$nowcol, 'E-MAIL (電郵): '.$potem["remark"]["c3"]);
//$sheet->setCellValue('I'.$nowcol, $potem["remark"]["c4"]);
////$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
$sheet->setCellValue('O'.$nowcol, $potem["remark"]["c3"]);
$nowcol++;

$sheet->setCellValue('O'.$nowcol, $potem["remark"]["c4"]);

$sheet->getPageSetup()
    ->setOrientation(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_PORTRAIT);  //竖放置
$sheet->getPageSetup()
    ->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);  //A4
$sheet->getPageSetup()->setFitToPage(true); //将工作表调整为一页

unset($_SESSION['potem30'] ); //注销SESSION

$filenameout = 'PO_'.$potem['shortName'].'_'.$potem['pono'];
outExcel($spreadsheet, $filenameout);


