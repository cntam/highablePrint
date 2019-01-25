<?php
/* 樂友膠袋廠*/
require_once 'aidenfunc.php';
$potem4 =  $_SESSION['potem4'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/potem4.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle("sheet1");
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
//$sheet->getDefaultColumnDimension()->setWidth(20);  //设置默认列宽
$sheet->getColumnDimension('A')->setWidth(20);  //列宽度
$sheet->getColumnDimension('B')->setWidth(40);  //列宽度
$sheet->getColumnDimension('C')->setWidth(25);  //列宽度
$sheet->getColumnDimension('D')->setWidth(35);  //列宽度
$sheet->getColumnDimension('E')->setWidth(16);  //列宽度
$sheet->getColumnDimension('F')->setWidth(16);  //列宽度

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
$styleArrayr = [

    'borders' => [

        'right' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],

    ],

];

$styleArraybu = [

    'borders' => [

        'bottom' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],

    ],

];

//填数据
$sheet->setCellValue('B5', $potem4["tosb"]);
$sheet->setCellValue('D6', $potem4 ["podate"]);
$sheet->setCellValue('B6', $potem4["toaddr"]["a1"]);
$sheet->setCellValue('B7', $potem4["toaddr"]["a2"]);
$sheet->setCellValue('B8', $potem4["toaddr"]["a3"]);
$sheet->setCellValue('D8', $potem4["toaddr"]["a4"]);
$sheet->setCellValue('B9', $potem4["toaddr"]["a5"]);
$sheet->setCellValue('D9', $potem4["toaddr"]["a6"]);


//中部form
$sheet->setCellValue('B11', $potem4["toaddr"]["a7"]);
$sheet->setCellValue('B12', $potem4["orderform"]["midpono"]);

$nowcol = 14;
//$spreadsheet->getActiveSheet()->mergeCells("A{$nowcol}:F{$nowcol}");
//$spreadsheet->getActiveSheet()->setCellValue('A'.$nowcol, '(PO NO:  '.$potem4["orderform"]["midpono"].' 注：請在開發票時把"PO NO"寫上，不可重複)');
////$spreadsheet->getActiveSheet()->setCellValue('I'.$nowcol, $potem4["invoiceform"]["amout"]);
////
//$nowcol++;
//$nowcol++;

for($x = 0 ,$c = 1; $x <= $potem4["orderform"]["formnum"]; $x++ ,$c++){

$f19 = 14 + 1 * $x;

//$spreadsheet->getActiveSheet()->mergeCells("B{$f19}:E{$f19}");

$formarr = array('A'.$f19,'B'.$f19,'C'.$f19,'D'.$f19);

    for($i = 1,$y = 0; $i <= $potem4["orderform"]["brrnum"] ; $i++ ,$y++){

        $sheet->setCellValue($formarr[$y],  $potem4["orderform"]['b'.$i][$x]);

    }


    $nowcol = 14  +   1 * $c;

//    $spreadsheet->getActiveSheet()->getStyle('A'.$f19)->applyFromArray($styleArray1);
//    $spreadsheet->getActiveSheet()->getStyle("B{$f19}:E{$f19}")->applyFromArray($styleArray1);
//    $spreadsheet->getActiveSheet()->getStyle('F'.$f19)->applyFromArray($styleArray1);
//    $spreadsheet->getActiveSheet()->getStyle('G'.$f19)->applyFromArray($styleArray1);

    if($x >12){
        $sheet->insertNewRowBefore($nowcol, 1);
    }

}



//底部REMARK
$nowcol = $potem4["orderform"]["formnum"] > 12 ? ($nowcol + 2) : 29;
//$sheet->getCell('A1')->setValue($nowcol); 貨送以下地址
//$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
//$sheet->setCellValue('A'.$nowcol, '貨送以下地址');
//$nowcol++;
//
//$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
$sheet->setCellValue('B'.$nowcol, $potem4["remark"]["c1"]);
$nowcol++;

//
//$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
$sheet->setCellValue('B'.$nowcol, $potem4["remark"]["c2"]);
$nowcol++;

//$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
$sheet->setCellValue('A'.$nowcol, $potem4["remark"]["c3"]);
$nowcol++;
$sheet->setCellValue('B'.$nowcol, $potem4["remark"]["c4"]);
$nowcol++;

$sheet->getPageSetup()->setFitToPage(true); //将工作表调整为一页

unset($_SESSION['potem4'] ); //注销SESSION

$filenameout = 'PO_'.$potem4['shortName'];
outExcel($spreadsheet,$filenameout);

