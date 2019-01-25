<?php
require_once 'aidenfunc.php';
header("Content-type: text/html; charset=utf-8");


$potem15 =  $_SESSION['potem15'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/potem15.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$spreadsheet->getActiveSheet()->setTitle("sheet1");
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
//$spreadsheet->getActiveSheet()->getDefaultColumnDimension()->setWidth(30);  //设置默认列宽
//for($j=0;$j<=6;$j++){
//    $col = chr(65 + $j);
//    $spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(20);  //列宽度
//}

$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(25);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(15);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(10);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(15);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('H')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('I')->setWidth(30);  //列宽度

$spreadsheet->getDefaultStyle()->getFont()->setSize(7);

$styleArray1 = [
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
        'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
        'wrapText' => true,
        'ShrinkToFit'=>true,
    ],
    'font' => [
        'Size' => '6',
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

    ]

];
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
$spreadsheet->getActiveSheet()->setCellValue('B6', $potem15["tosb"]);
$spreadsheet->getActiveSheet()->setCellValue('B8', $potem15["podate"]);
////
//$spreadsheet->getActiveSheet()->setCellValue('G8', $potem15["toaddr"]["a1"]);
//$spreadsheet->getActiveSheet()->setCellValue('B9', $potem15["toaddr"]["a2"]);
//$spreadsheet->getActiveSheet()->setCellValue('B11', $potem15["toaddr"]["a3"]);
//$spreadsheet->getActiveSheet()->setCellValue('B12', $potem15["toaddr"]["a4"]);
$toaddr = array('H6','B7','H7','H8','B9','H9','B12','B13','B14','B15','B16','B17','B18','B19','B20','G12','G13','G14','G15','G16','G17','G18','G19','G20');  //

for($i = 1,$y = 0; $i <= count($toaddr) ; $i++ ,$y++){

    $sheet->setCellValue($toaddr[$y],  $potem15["toaddr"]["a".$i]);

}
//
//
////中部form

//$nowcol = 24;
//////$spreadsheet->getActiveSheet()->mergeCells("A{$nowcol}:F{$nowcol}");
//$spreadsheet->getActiveSheet()->setCellValue('B'.$nowcol, $potem15["orderform"]["midpono"]);
//////$spreadsheet->getActiveSheet()->setCellValue('I'.$nowcol, $potem15["invoiceform"]["amout"]);
//
//
for($x = 0 ,$c = 1; $c <= $potem15["orderform"]["formnum"]; $x++ ,$c++){

$f19 = 24 + 1 * $x;

//    $spreadsheet->getActiveSheet()->mergeCells("A{$f19}:B{$f19}");
//    $spreadsheet->getActiveSheet()->mergeCells("C{$f19}:G{$f19}");


$formarr = array('A'.$f19,'B'.$f19,'C'.$f19,'D'.$f19);

    for($i = 1,$y = 0; $i <= $potem15["orderform"]["brrnum"] ; $i++ ,$y++){

        $sheet->setCellValue($formarr[$y],  $potem15["orderform"]['b'.$i][$x]);

    }

    $nowcol = 24  +  1 * $c;


    if($x >15){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($nowcol, 1);
    }

}
//$nowcol = $potem15["orderform"]["formnum"] > 4 ? ($nowcol + 1) : 21;
//$spreadsheet->getActiveSheet()->setCellValue('C'.$nowcol, $potem15["toaddr"]["a5"]);
//$nowcol++;
//
//$spreadsheet->getActiveSheet()->setCellValue('C'.$nowcol, $potem15["toaddr"]["a6"]);
//$nowcol++;
//$spreadsheet->getActiveSheet()->setCellValue('C'.$nowcol, $potem15["toaddr"]["a7"]);
//$nowcol++;
//$spreadsheet->getActiveSheet()->setCellValue('A'.$nowcol, $potem15["toaddr"]["a8"]);
//$nowcol++;
//$nowcol++;
//$nowcol++;
//$nowcol++;
//$nowcol++;
//$nowcol++;
//$nowcol++;
//
//$spreadsheet->getActiveSheet()->setCellValue('C'.$nowcol, $potem15["remark"]["c1"]);
//$nowcol++;
//$spreadsheet->getActiveSheet()->setCellValue('C'.$nowcol, $potem15["remark"]["c2"]);

////
////$spreadsheet->getActiveSheet()->setCellValue('B'.$nowcol, 'E-MAIL (電郵): '.$potem15["remark"]["c3"]);
////$spreadsheet->getActiveSheet()->setCellValue('I'.$nowcol, $potem15["remark"]["c4"]);
//////$spreadsheet->getActiveSheet()->mergeCells("A{$nowcol}:E{$nowcol}");
//$spreadsheet->getActiveSheet()->setCellValue('O'.$nowcol, $potem15["remark"]["c3"]);
//$nowcol++;
//
//$spreadsheet->getActiveSheet()->setCellValue('O'.$nowcol, $potem15["remark"]["c4"]);

$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

unset($_SESSION['potem15'] ); //注销SESSION

$filenameout = 'PO_'.$potem15['shortName'];
outExcel($spreadsheet,$filenameout);

