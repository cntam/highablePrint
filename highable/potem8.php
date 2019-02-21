<?php
require_once 'aidenfunc.php';
$pot =  $_SESSION['potem8'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/potem8.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle("sheet1");
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
//$sheet->getDefaultColumnDimension()->setWidth(20);  //设置默认列宽
$sheet->getColumnDimension('A')->setWidth(25);  //列宽度
$sheet->getColumnDimension('B')->setWidth(55);  //列宽度
$sheet->getColumnDimension('C')->setWidth(30);  //列宽度
$sheet->getColumnDimension('D')->setWidth(15);  //列宽度
$sheet->getColumnDimension('E')->setWidth(35);  //列宽度
//$sheet->getColumnDimension('F')->setWidth(25);  //列宽度
//$sheet->getColumnDimension('G')->setWidth(16);  //列宽度
//$sheet->getColumnDimension('H')->setWidth(15);  //列宽度
//$sheet->getColumnDimension('I')->setWidth(15);  //列宽度
$spreadsheet->getDefaultStyle()->getFont()->setSize(7);

$styleArray2 = [
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
        'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
        'wrapText' => true,
        'ShrinkToFit'=>true,
    ],
    'font' => [
        'Size' => '5',
    ]

];

//填数据
//poheader
setCell($sheet, "A1", $pot["remark"]["poheader"]["poheada1"], $noborderCenter);
setCell($sheet, "A2", $pot["remark"]["poheader"]["poheada2"].' '.$pot["remark"]["poheader"]["poheada3"], $Size12noborderCenter);
//setCell($sheet, "A4", $potem6["remark"]["poheader"]["poheada3"], $noborderCenter);
setCell($sheet, "A3", $pot["remark"]["poheader"]["poheada4"], $noborderCenter);
//setCell($sheet, "A6", $potem6["remark"]["poheader"]["poheada6"], $noborderCenter);

$sheet->setCellValue('B5', $pot["tosb"]);
$sheet->getStyle('B5')->applyFromArray($styleArray2);
$sheet->setCellValue('E6', $pot["podate"]);
$sheet->setCellValue('B6', $pot["toaddr"]["a1"]);
$sheet->setCellValue('B10', $pot["toaddr"]["a2"]);
$sheet->setCellValue('B7', $pot["toaddr"]["a3"]);
$sheet->setCellValue('B8', $pot["toaddr"]["a4"]);
$sheet->setCellValue('B9', $pot["toaddr"]["a5"]);
$sheet->setCellValue('E9', $pot["toaddr"]["a6"]);


//中部form

$nowcol = 11;
//$sheet->mergeCells("A{$nowcol}:F{$nowcol}");
$sheet->mergeCells("A11:E11");
$sheet->setCellValue('A'.$nowcol, '(PO NO.'.$pot["orderform"]["midpono"].'請在發票上寫上制單號及注明PO NO.,不可重復,謝)');
//$sheet->setCellValue('I'.$nowcol, $pot["invoiceform"]["amout"]);


for($x = 0 ,$c = 1; $c <= $pot["orderform"]["formnum"]; $x++ ,$c++){

$f19 = 19 + 1 * $x;

//$sheet->mergeCells("B{$f19}:E{$f19}");

$formarr = array('A'.$f19,'B'.$f19,'C'.$f19,'D'.$f19,'E'.$f19);

    for($i = 1,$y = 0; $i <= $pot["orderform"]["brrnum"] ; $i++ ,$y++){

//        $sheet->setCellValue($formarr[$y],  $pot["orderform"]['b'.$i][$x]);
        setCell($sheet, $formarr[$y], $pot["orderform"]['b'.$i][$x], $noborderCenter);

    }


    $nowcol = 19  +   1 * $c;


    if($x >3){
        $sheet->insertNewRowBefore($nowcol, 1);
    }

}
//$nowcol = $pot["orderform"]["formnum"] > 14 ? ($nowcol + 1) : 26;
////$sheet->getCell('A1')->setValue($nowcol); 貨送以下地址
////$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
////$sheet->setCellValue('A'.$nowcol, '貨送以下地址');
////$nowcol++;
//
//$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
$sheet->setCellValue('A12', $pot["remark"]["c1"]);


//$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
$sheet->setCellValue('A13', $pot["remark"]["c2"]);

////$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
//$sheet->setCellValue('A'.$nowcol, $pot["remark"]["c3"]);
//

$sheet->getPageSetup()
    ->setOrientation(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_PORTRAIT);  //竖放置
$sheet->getPageSetup()
    ->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);  //A4
$sheet->getPageSetup()->setFitToPage(true); //将工作表调整为一页

unset($_SESSION['potem8'] ); //注销SESSION

$filenameout = 'PO_'.$pot['pono'];
outExcel($spreadsheet,$filenameout);

