<?php
require_once 'aidenfunc.php';
$potem14 =  $_SESSION['potem14'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/potem14.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle("sheet1");
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
//$sheet->getDefaultColumnDimension()->setWidth(30);  //设置默认列宽
for($j=0;$j<=6;$j++){
    $col = chr(65 + $j);
    $sheet->getColumnDimension($col)->setWidth(20);  //列宽度
}

//$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(25);  //列宽度
//$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(5);  //列宽度
//$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(15);  //列宽度

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
$sheet->setCellValue('B8', $potem14["tosb"]);
$sheet->setCellValue('B10', $potem14["podate"]);
////
$sheet->setCellValue('G8', $potem14["toaddr"]["a1"]);
$sheet->setCellValue('B9', $potem14["toaddr"]["a2"]);
$sheet->setCellValue('B11', $potem14["toaddr"]["a3"]);
$sheet->setCellValue('B12', $potem14["toaddr"]["a4"]);
//$toaddr = array('Z9','B11','T11','B13','T13','F15','O15','Y15','B23','B34','U23','B41','K41','U41','Z41','AE41','J43','AC43');  //,'C12','D12','E12','F12','G12','H12','I12','J13','K12','K13','J5'
//
//for($i = 1,$y = 0; $i <= count($toaddr) ; $i++ ,$y++){
//
//    $sheet->setCellValue($toaddr[$y],  $potem14["toaddr"]["a".$i]);
//
//}
//
//
////中部form
//
$nowcol = 14;
//////$sheet->mergeCells("A{$nowcol}:F{$nowcol}");
$sheet->setCellValue('B'.$nowcol, $potem14["toaddr"]["a9"]);
//////$sheet->setCellValue('I'.$nowcol, $potem14["invoiceform"]["amout"]);
//
//
for($x = 0 ,$c = 1; $c <= $potem14["orderform"]["formnum"]; $x++ ,$c++){

$f19 = 15 + 1 * $x;

    $sheet->mergeCells("A{$f19}:B{$f19}");
    $sheet->mergeCells("C{$f19}:G{$f19}");


$formarr = array('A'.$f19,'C'.$f19);

    for($i = 1,$y = 0; $i <= $potem14["orderform"]["brrnum"] ; $i++ ,$y++){

        $sheet->setCellValue($formarr[$y],  $potem14["orderform"]['b'.$i][$x]);

    }

    $nowcol = 15  +  1 * $c;


    if($x >4){
        $sheet->insertNewRowBefore($nowcol, 1);
    }

}
$nowcol = $potem14["orderform"]["formnum"] > 4 ? ($nowcol + 1) : 21;
$sheet->setCellValue('C'.$nowcol, $potem14["toaddr"]["a5"]);
$nowcol++;

$sheet->setCellValue('C'.$nowcol, $potem14["toaddr"]["a6"]);
$nowcol++;
$sheet->setCellValue('C'.$nowcol, $potem14["toaddr"]["a7"]);
$nowcol++;
$sheet->setCellValue('A'.$nowcol, $potem14["toaddr"]["a8"]);
$nowcol++;
$nowcol++;
$nowcol++;
$nowcol++;
$nowcol++;
$nowcol++;
$nowcol++;

$sheet->setCellValue('C'.$nowcol, $potem14["remark"]["c1"]);
$nowcol++;
$sheet->setCellValue('C'.$nowcol, $potem14["remark"]["c2"]);

////
////$spreadsheet->getActiveSheet()->setCellValue('B'.$nowcol, 'E-MAIL (電郵): '.$potem14["remark"]["c3"]);
////$spreadsheet->getActiveSheet()->setCellValue('I'.$nowcol, $potem14["remark"]["c4"]);
//////$spreadsheet->getActiveSheet()->mergeCells("A{$nowcol}:E{$nowcol}");
//$spreadsheet->getActiveSheet()->setCellValue('O'.$nowcol, $potem14["remark"]["c3"]);
//$nowcol++;
//
//$sheet->setCellValue('O'.$nowcol, $potem14["remark"]["c4"]);

$sheet->getPageSetup()->setFitToPage(true); //将工作表调整为一页

unset($_SESSION['potem14'] ); //注销SESSION

$filenameout = 'PO_'.$potem14['shortName'];
outExcel($spreadsheet,$filenameout);
