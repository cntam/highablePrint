<?php
require_once 'aidenfunc.php';
$potem7 =  $_SESSION['potem7'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/potem7.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle("sheet1");
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
//$sheet->getDefaultColumnDimension()->setWidth(20);  //设置默认列宽
$sheet->getColumnDimension('A')->setWidth(25);  //列宽度
$sheet->getColumnDimension('B')->setWidth(25);  //列宽度
$sheet->getColumnDimension('C')->setWidth(25);  //列宽度
$sheet->getColumnDimension('D')->setWidth(25);  //列宽度
$sheet->getColumnDimension('E')->setWidth(25);  //列宽度
$sheet->getColumnDimension('F')->setWidth(25);  //列宽度
//$sheet->getColumnDimension('G')->setWidth(16);  //列宽度
//$sheet->getColumnDimension('H')->setWidth(15);  //列宽度
//$sheet->getColumnDimension('I')->setWidth(15);  //列宽度
$spreadsheet->getDefaultStyle()->getFont()->setSize(7);

//填数据
// poheader
setCell($sheet, "A1", $potem7["remark"]["poheader"]["poheada1"], $noborderCenter);

$sheet->setCellValue('B3', $potem7["tosb"]);
$sheet->setCellValue('E3', $potem7 ["podate"]);
$sheet->setCellValue('B4', $potem7["toaddr"]["a1"]);
//$sheet->setCellValue('E4', $potem7["toaddr"]["a2"]);
setCell($sheet, "E4", $potem7["toaddr"]["a2"], $noborderLeft);
$sheet->setCellValue('B5', $potem7["toaddr"]["a3"]);
$sheet->setCellValue('E5', $potem7["toaddr"]["a4"]);
$sheet->setCellValue('B6', $potem7["toaddr"]["a5"]);
$sheet->setCellValue('E6', $potem7["toaddr"]["a6"]);


//中部form

$nowcol = 8;
//$sheet->mergeCells("A{$nowcol}:F{$nowcol}");
$sheet->setCellValue('A'.$nowcol, '(PO NO:  '.$potem7["orderform"]["midpono"].' （注：请在开月结单时把“PO NO”写上，不可重复，并且写上制单号）');
//$sheet->setCellValue('I'.$nowcol, $potem7["invoiceform"]["amout"]);
//
//$nowcol++;
$nowcol++;

for($x = 0 ,$c = 1; $x <= $potem7["orderform"]["formnum"]; $x++ ,$c++){

$f19 = 10 + 1 * $x;

//$sheet->mergeCells("B{$f19}:E{$f19}");

$formarr = array('A'.$f19,'B'.$f19,'C'.$f19,'D'.$f19,'E'.$f19,'F'.$f19);

    for($i = 1,$y = 0; $i <= $potem7["orderform"]["brrnum"] ; $i++ ,$y++){

        $sheet->setCellValue($formarr[$y],  $potem7["orderform"]['b'.$i][$x]);

    }


    $nowcol = 10  +   1 * $c;



    if($x >14){
        $sheet->insertNewRowBefore($nowcol, 1);
    }

}
$nowcol = $potem7["orderform"]["formnum"] > 14 ? ($nowcol + 1) : 26;
//$sheet->getCell('A1')->setValue($nowcol); 貨送以下地址
//$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
//$sheet->setCellValue('A'.$nowcol, '貨送以下地址');
//$nowcol++;

//$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
$sheet->setCellValue('B'.$nowcol, $potem7["remark"]["c1"]);
$nowcol++;
$nowcol++;
$nowcol++;
$nowcol++;

//$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
$sheet->setCellValue('D'.$nowcol, $potem7["remark"]["c2"]);
$nowcol++;

//$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
$sheet->setCellValue('A'.$nowcol, $potem7["remark"]["c3"]);



$sheet->getPageSetup()
    ->setOrientation(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_PORTRAIT);  //竖放置
$sheet->getPageSetup()
    ->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);  //A4
$sheet->getPageSetup()->setFitToPage(true); //将工作表调整为一页

unset($_SESSION['potem7'] ); //注销SESSION

$filenameout = 'PO_'.$potem7['pono'];
outExcel($spreadsheet,$filenameout);

