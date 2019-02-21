<?php
require_once 'aidenfunc.php';
$potem10 =  $_SESSION['potem10'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/potem10.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle("sheet1");
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
//$sheet->getDefaultColumnDimension()->setWidth(20);  //设置默认列宽
$sheet->getColumnDimension('A')->setWidth(25);  //列宽度
$sheet->getColumnDimension('B')->setWidth(20);  //列宽度
$sheet->getColumnDimension('C')->setWidth(20);  //列宽度
$sheet->getColumnDimension('D')->setWidth(20);  //列宽度
$sheet->getColumnDimension('E')->setWidth(20);  //列宽度
$sheet->getColumnDimension('F')->setWidth(20);  //列宽度
$sheet->getColumnDimension('G')->setWidth(20);  //列宽度
$sheet->getColumnDimension('H')->setWidth(20);  //列宽度
$sheet->getColumnDimension('I')->setWidth(20);  //列宽度
$sheet->getColumnDimension('J')->setWidth(20);  //列宽度
$sheet->getColumnDimension('K')->setWidth(20);  //列宽度
$spreadsheet->getDefaultStyle()->getFont()->setSize(7);

//填数据
$sheet->setCellValue('K3', '共    1   页');
$sheet->setCellValue('K4', '第   1  页');

$sheet->setCellValue('A11', '色号');
$sheet->setCellValue('B11', '色号');

$sheet->setCellValue('B7', $potem10["tosb"]);
//$sheet->setCellValue('F7', $potem10["podate"]);


$toaddr = array('B6','J5','A10','B10','E10','F10','H10','I10','J10','K10','B12','C12','D12','E12','F12','G12','H12','I12','J13','K12','K13');

    for($i = 1,$y = 0; $i <= count($toaddr) ; $i++ ,$y++){

        $sheet->setCellValue($toaddr[$y],  $potem10["toaddr"]["a".$i]);

    }


//中部form

$nowcol = 13;
//$sheet->mergeCells("A{$nowcol}:F{$nowcol}");
//$sheet->setCellValue('A'.$nowcol, 'PO NO: '.$potem10["orderform"]["midpono"].'   注：請在開發票時把“PONO”寫上，不可重復，并且寫上制單號）');
//$sheet->setCellValue('I'.$nowcol, $potem10["invoiceform"]["amout"]);


for($x = 0 ,$c = 1; $c <= $potem10["orderform"]["formnum"]; $x++ ,$c++){

$f19 = 13 + 1 * $x;

//$sheet->mergeCells("B{$f19}:E{$f19}");

$formarr = array('A'.$f19,'B'.$f19,'C'.$f19,'D'.$f19,'E'.$f19,'F'.$f19,'G'.$f19,'H'.$f19,'I'.$f19);

    for($i = 1,$y = 0; $i <= $potem10["orderform"]["brrnum"] ; $i++ ,$y++){

        $sheet->setCellValue($formarr[$y],  $potem10["orderform"]['b'.$i][$x]);

    }

    $nowcol = 13  +   1 * $c;


    if($c >3){
        $sheet->insertNewRowBefore($nowcol, 1);
    }

}
$nowcol = $potem10["orderform"]["formnum"] > 3 ? ($nowcol + 2) : 18;
//$sheet->getCell('A1')->setValue($nowcol); 貨送以下地址
//$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
//$sheet->setCellValue('A'.$nowcol, '貨送以下地址');
//$nowcol++;

//$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
$sheet->setCellValue('A'.$nowcol, '送货地址：'.$potem10["remark"]["c1"]);


//$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
$sheet->setCellValue('F'.$nowcol, $potem10["remark"]["c2"]);

//$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
$sheet->setCellValue('J'.$nowcol, $potem10["remark"]["c3"]);
$nowcol++;
$nowcol++;
$sheet->setCellValue('B'.$nowcol, $potem10["remark"]["c4"]);
$sheet->setCellValue('J'.$nowcol, $potem10["remark"]["c5"]);

//$sheet->getPageSetup()
//    ->setOrientation(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_PORTRAIT);  //竖放置
$sheet->getPageSetup()
    ->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);  //A4
$sheet->getPageSetup()->setFitToPage(true); //将工作表调整为一页

unset($_SESSION['potem10'] ); //注销SESSION

$filenameout = 'PO_'.$potem10['pono'];
outExcel($spreadsheet,$filenameout);

