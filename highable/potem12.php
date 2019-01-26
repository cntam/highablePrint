<?php
require_once 'aidenfunc.php';

$potem12 =  $_SESSION['potem12'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/potem12.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle("sheet1");
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
//$sheet->getDefaultColumnDimension()->setWidth(20);  //设置默认列宽
$sheet->getColumnDimension('A')->setWidth(16);  //列宽度
$sheet->getColumnDimension('B')->setWidth(16);  //列宽度
$sheet->getColumnDimension('C')->setWidth(16);  //列宽度
$sheet->getColumnDimension('D')->setWidth(16);  //列宽度
$sheet->getColumnDimension('E')->setWidth(16);  //列宽度
$sheet->getColumnDimension('F')->setWidth(16);  //列宽度
$sheet->getColumnDimension('G')->setWidth(16);  //列宽度
$sheet->getColumnDimension('H')->setWidth(16);  //列宽度
$sheet->getColumnDimension('I')->setWidth(26);  //列宽度
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
$sheet->setCellValue('A7', 'TO: '.$potem12["tosb"]);
$sheet->setCellValue('H6', 'DATE: '.$potem12["podate"]);

$toaddr = array('A8','H7','A9','H8','A10','F10','A11','B12','B13','B14','C16','H16','C18','H18','C20','H20','C23','H23','C25','H25');  //,'C12','D12','E12','F12','G12','H12','I12','J13','K12','K13','J5'

for($i = 1,$y = 0; $i <= count($toaddr) ; $i++ ,$y++){

    $sheet->setCellValue($toaddr[$y],  $potem12["toaddr"]["a".$i]);

}

//中部form

$nowcol = 28;
//$sheet->mergeCells("A{$nowcol}:F{$nowcol}");
$sheet->setCellValue('E'.$nowcol, $potem12["orderform"]["midpono"]);
////$sheet->setCellValue('I'.$nowcol, $potem12["invoiceform"]["amout"]);


for($x = 0 ,$c = 1; $c <= $potem12["orderform"]["formnum"]; $x++ ,$c++){

$f19 = 31 + 1 * $x;

//$spreadsheet->getActiveSheet()->mergeCells("B{$f19}:C{$f19}");

$formarr = array('A'.$f19,'B'.$f19,'C'.$f19,'D'.$f19,'E'.$f19,'F'.$f19,'G'.$f19,'H'.$f19,'I'.$f19);

    for($i = 1,$y = 0; $i <= $potem12["orderform"]["brrnum"] ; $i++ ,$y++){

        $sheet->setCellValue($formarr[$y],  $potem12["orderform"]['b'.$i][$x]);

    }

    $nowcol = 31  +   1 * $c;


    if($x >1){
        $sheet->insertNewRowBefore($nowcol, 1);
    }

}
$nowcol = $potem12["orderform"]["formnum"] > 1 ? ($nowcol + 2) : 35;

$sheet->setCellValue('E'.$nowcol,  $potem12["toaddr"]["a48"]);


$nowcol2 = $potem12["orderform"]["formnum"] > 1 ? ($nowcol+3) : 38;

//$nowcol3 = $potem12["orderform"]["formnum"] > 1 ? ($nowcol + 6) : 42;
//$nowcol4 = $potem12["orderform"]["formnum"] > 1 ? ($nowcol + 5) : 43;

for($t=0,$r = 1;$t < 3;$t++ ,$r++){
    $toaddr = array('A'.$nowcol2,'B'.$nowcol2,'C'.$nowcol2,'D'.$nowcol2,'E'.$nowcol2,'F'.$nowcol2,'G'.$nowcol2,'H'.$nowcol2,'I'.$nowcol2);

    for($i = 21 + (count($toaddr) * $t) ,$y = 0; $i < 21 + (count($toaddr) * $r)  ; $i++ ,$y++){

        $sheet->setCellValue($toaddr[$y],  $potem12["toaddr"]["a".$i]);

    }
    $nowcol2++;
}





//////$spreadsheet->getActiveSheet()->getCell('A1')->setValue($nowcol); 貨送以下地址
//////$spreadsheet->getActiveSheet()->mergeCells("A{$nowcol}:E{$nowcol}");
//////$spreadsheet->getActiveSheet()->setCellValue('A'.$nowcol, '貨送以下地址');
//////$nowcol++;
////
////$spreadsheet->getActiveSheet()->mergeCells("A{$nowcol}:E{$nowcol}");
//$spreadsheet->getActiveSheet()->setCellValue('D'.$nowcol, $potem12["toaddr"]["a20"]);
//$nowcol++;
//$nowcol++;
//$nowcol++;
//
//////$spreadsheet->getActiveSheet()->mergeCells("A{$nowcol}:E{$nowcol}");
//$spreadsheet->getActiveSheet()->setCellValue('D'.$nowcol, $potem12["toaddr"]["a21"]);
//$nowcol++;
//$nowcol++;
//$nowcol++;
//$nowcol++;
//
//
//
//$spreadsheet->getActiveSheet()->setCellValue('B'.$nowcol, 'ORDERED BY (經手人) :'.$potem12["remark"]["c1"]);
//$spreadsheet->getActiveSheet()->setCellValue('I'.$nowcol, $potem12["remark"]["c2"]);
//$nowcol++;
//$nowcol++;
//
//$spreadsheet->getActiveSheet()->setCellValue('B'.$nowcol, 'E-MAIL (電郵): '.$potem12["remark"]["c3"]);
//$spreadsheet->getActiveSheet()->setCellValue('I'.$nowcol, $potem12["remark"]["c4"]);
////$spreadsheet->getActiveSheet()->mergeCells("A{$nowcol}:E{$nowcol}");
//$spreadsheet->getActiveSheet()->setCellValue('E'.$nowcol, $potem12["remark"]["c3"]);
//$nowcol++;
//
//$spreadsheet->getActiveSheet()->setCellValue('A'.$nowcol, $potem12["remark"]["c4"]);

$sheet->getPageSetup()->setFitToPage(true); //将工作表调整为一页
//
//unset($_SESSION['potem12'] ); //注销SESSION

$filenameout = 'PO_'.$potem12['shortName'];
outExcel($spreadsheet,$filenameout);


