<?php
require_once 'aidenfunc.php';
$potem9 =  $_SESSION['potem9'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/potem9.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle("sheet1");
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
//$sheet->getDefaultColumnDimension()->setWidth(20);  //设置默认列宽
$sheet->getColumnDimension('A')->setWidth(25);  //列宽度
$sheet->getColumnDimension('B')->setWidth(30);  //列宽度
$sheet->getColumnDimension('C')->setWidth(25);  //列宽度
$sheet->getColumnDimension('D')->setWidth(25);  //列宽度
$sheet->getColumnDimension('E')->setWidth(20);  //列宽度
$sheet->getColumnDimension('F')->setWidth(15);  //列宽度
$sheet->getColumnDimension('G')->setWidth(30);  //列宽度
//$sheet->getColumnDimension('H')->setWidth(15);  //列宽度
//$sheet->getColumnDimension('I')->setWidth(15);  //列宽度
$spreadsheet->getDefaultStyle()->getFont()->setSize(7);


//填数据
//poheader
setCell($sheet, "A1", $potem9["remark"]["poheader"]["poheada1"], $noborderCenter);
setCell($sheet, "A2", $potem9["remark"]["poheader"]["poheada2"], $noborderCenter);
setCell($sheet, "A3", $potem9["remark"]["poheader"]["poheada3"], $noborderCenter);
setCell($sheet, "A4", $potem9["remark"]["poheader"]["poheada4"], $noborderCenter);
//setCell($sheet, "A5", $potem9["remark"]["poheader"]["poheada6"], $noborderCenter);

$sheet->setCellValue('B7', $potem9["tosb"]);
$sheet->setCellValue('F7', $potem9["podate"]);
$sheet->setCellValue('B8', $potem9["toaddr"]["a1"]);
$sheet->setCellValue('F8', $potem9["toaddr"]["a2"]);
$sheet->setCellValue('B9', $potem9["toaddr"]["a3"]);
$sheet->setCellValue('F9', $potem9["toaddr"]["a4"]);
$sheet->setCellValue('B10', $potem9["toaddr"]["a5"]);
$sheet->setCellValue('F10', $potem9["toaddr"]["a6"]);


//中部form

$nowcol = 12;
//$sheet->mergeCells("A{$nowcol}:F{$nowcol}");
$sheet->setCellValue('A'.$nowcol, 'PO NO: '.$potem9["orderform"]["midpono"].'   注：請在開發票時把“PONO”寫上，不可重復，并且寫上制單號）');
//$sheet->setCellValue('I'.$nowcol, $potem9["invoiceform"]["amout"]);


for($x = 0 ,$c = 1; $c <= $potem9["orderform"]["formnum"]; $x++ ,$c++){

$f19 = 14 + 1 * $x;

//$sheet->mergeCells("B{$f19}:E{$f19}");

$formarr = array('A'.$f19,'B'.$f19,'C'.$f19,'D'.$f19,'E'.$f19,'F'.$f19,'G'.$f19);

    for($i = 1,$y = 0; $i <= $potem9["orderform"]["brrnum"] ; $i++ ,$y++){

        $sheet->setCellValue($formarr[$y],  $potem9["orderform"]['b'.$i][$x]);

    }

    $nowcol = 14  +   1 * $c;


    if($x >2){
        $sheet->insertNewRowBefore($nowcol, 1);
    }

}
$nowcol = $potem9["orderform"]["formnum"] > 2 ? ($nowcol + 1) : 18;
//$sheet->getCell('A1')->setValue($nowcol); 貨送以下地址
//$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
//$sheet->setCellValue('A'.$nowcol, '貨送以下地址');
//$nowcol++;

//$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
$sheet->setCellValue('B'.$nowcol, '内地交货：送货地址：'.$potem9["remark"]["c1"]);
$nowcol++;

//$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
$sheet->setCellValue('B'.$nowcol, $potem9["remark"]["c2"]);
$nowcol++;
//$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
$sheet->setCellValue('E'.$nowcol, $potem9["remark"]["c3"]);
$nowcol++;

$sheet->setCellValue('A'.$nowcol, $potem9["remark"]["c4"]);

$sheet->getPageSetup()
    ->setOrientation(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_PORTRAIT);  //竖放置
$sheet->getPageSetup()
    ->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);  //A4
$sheet->getPageSetup()->setFitToPage(true); //将工作表调整为一页

//unset($_SESSION['potem9'] ); //注销SESSION

$filenameout = 'PO_'.$potem9['shortName'].'_'.$potem9['pono'];
outExcel($spreadsheet,$filenameout);
