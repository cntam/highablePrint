<?php
require_once 'aidenfunc.php';
$potem6 =  $_SESSION['potem6'];

//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/potem6.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle("sheet1");
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
//$sheet->getDefaultColumnDimension()->setWidth(20);  //设置默认列宽
$sheet->getColumnDimension('A')->setWidth(20);  //列宽度
$sheet->getColumnDimension('B')->setWidth(25);  //列宽度
$sheet->getColumnDimension('C')->setWidth(25);  //列宽度
$sheet->getColumnDimension('D')->setWidth(25);  //列宽度
$sheet->getColumnDimension('E')->setWidth(20);  //列宽度
$sheet->getColumnDimension('F')->setWidth(20);  //列宽度
$sheet->getColumnDimension('G')->setWidth(20);  //列宽度
//
//$sheet->getColumnDimension('I')->setWidth(15);  //列宽度
$spreadsheet->getDefaultStyle()->getFont()->setSize(7);


//填数据
//header
setCell($sheet, "A2", $potem6["remark"]["poheader"]["poheada1"], $noborderCenter);
setCell($sheet, "A3", '大陆地址:'.$potem6["remark"]["poheader"]["poheada2"].' '.$potem6["remark"]["poheader"]["poheada3"], $noborderCenter);
//setCell($sheet, "A4", $potem6["remark"]["poheader"]["poheada3"], $noborderCenter);
setCell($sheet, "A4", $potem6["remark"]["poheader"]["poheada4"].' '.$potem6["remark"]["poheader"]["poheada5"], $noborderCenter);
//setCell($sheet, "A6", $potem6["remark"]["poheader"]["poheada6"], $noborderCenter);

$sheet->setCellValue('B6', $potem6["tosb"]);
$sheet->setCellValue('F7', $potem6 ["podate"]);
$sheet->setCellValue('F6', $potem6["toaddr"]["a1"]);
$sheet->setCellValue('B7', $potem6["toaddr"]["a2"]);



//中部form
//$sheet->setCellValue('B11', $potem6["toaddr"]["a7"]);
//$sheet->setCellValue('B12', $potem6["orderform"]["midpono"]);

$nowcol = 11;
//$sheet->mergeCells("A{$nowcol}:F{$nowcol}");
//$sheet->setCellValue('A'.$nowcol, '(PO NO:  '.$potem6["orderform"]["midpono"].' 注：請在開發票時把"PO NO"寫上，不可重複)');
////$sheet->setCellValue('I'.$nowcol, $potem6["invoiceform"]["amout"]);
////
//$nowcol++;
//$nowcol++;
//
for($x = 0 ,$c = 1; $x < $potem6["orderform"]["formnum"]; $x++ ,$c++){

$f19 = 11 + 1 * $x;

$sheet->mergeCells("B{$f19}:D{$f19}");

$formarr = array('A'.$f19,'B'.$f19,'E'.$f19,'F'.$f19,'G'.$f19);

    for($i = 1,$y = 0; $i < $potem6["orderform"]["brrnum"] ; $i++ ,$y++){
        if ($i == 4){
            setCell($sheet, $formarr[$y], $potem6["orderform"]['b'.$i][$x]/100, $noborderCenter);
        }else{
            setCell($sheet, $formarr[$y], $potem6["orderform"]['b'.$i][$x], $noborderCenter);
        }
    }
    $nowcol = 11  +   1 * $c;

    if($x >9){
        $sheet->insertNewRowBefore($nowcol, 1);
    }

}


//底部REMARK
$nowcol = $potem6["orderform"]["formnum"] > 9? ($nowcol + 1) : 22;
//$sheet->getCell('A1')->setValue($nowcol);

$sheet->setCellValue('E'.$nowcol, $potem6["toaddr"]["a3"]);
//$sheet->setCellValue('F'.$nowcol, $potem6["toaddr"]["a4"]/100);
$sheet->setCellValue('G'.$nowcol, $potem6["toaddr"]["a5"]);
$nowcol++;

//$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
$sheet->setCellValue('A'.$nowcol, '備註：'.$potem6["remark"]["c1"]);
$nowcol++;
////
////$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
//$sheet->setCellValue('B'.$nowcol, $potem6["remark"]["c1"]);
$nowcol++;
$nowcol++;
//
////
//$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
$sheet->setCellValue('A'.$nowcol, $potem6["remark"]["c2"]);
$nowcol++;
//
////$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
$sheet->setCellValue('A'.$nowcol, 'FAX:'.$potem6["remark"]["c3"]);
$nowcol++;
$sheet->setCellValue('A'.$nowcol, 'E-mail:'.$potem6["remark"]["c4"]);
$nowcol++;
$sheet->setCellValue('A'.$nowcol, '交貨期:'.$potem6["remark"]["c5"]);


$sheet->getPageSetup()
    ->setOrientation(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_PORTRAIT);  //竖放置
$sheet->getPageSetup()
    ->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);  //A4
$sheet->getPageSetup()->setFitToPage(true); //将工作表调整为一页

//unset($_SESSION['potem6'] ); //注销SESSION

$filenameout = 'PO_'.$potem6['shortName'].'_'.$potem6['pono'];
outExcel($spreadsheet, $filenameout);

