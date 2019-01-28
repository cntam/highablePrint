<?php

require_once 'aidenfunc.php';

$potem =  $_SESSION['potem3'];

//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/potem3.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle("sheet1");
$spreadsheet->getDefaultStyle()->getFont()->setName('SimSun');
//$sheet->getDefaultColumnDimension()->setWidth(30);  //设置默认列宽
for($j=0;$j<=8;$j++){
    $col = chr(65 + $j);
    $sheet->getColumnDimension($col)->setWidth(24);  //列宽度
}

//$sheet->getColumnDimension('A')->setWidth(25);  //列宽度
//$sheet->getColumnDimension('B')->setWidth(5);  //列宽度
//$sheet->getColumnDimension('C')->setWidth(15);  //列宽度

$spreadsheet->getDefaultStyle()->getFont()->setSize(7);


//填数据
//header
setCell($sheet,"B8", $potem["tosb"], $noborderLeft);
setCell($sheet,"A9", $potem["toaddr"]['a1'], $noborderLeft);
setCell($sheet,"B10", $potem["toaddr"]['a2'], $noborderLeft);
setCell($sheet,"B11", $potem["toaddr"]['a3'], $noborderLeft);
setCell($sheet,"B12", $potem["toaddr"]['a4'], $noborderLeft);

setCell($sheet,"G8", $potem["podate"], $noborderLeft);
setCell($sheet,"F12", $potem["toaddr"]['a5'], $noborderLeft);

//表格动态
if ($potem["orderform"]["brrnum"] > 0) {
//    $col = 'A';
    $arr = array('A', 'B', 'C', 'D', 'F', 'G');
    for ($a = 0, $b = 1; $a < 6; $a++, $b++) {
        $row = 17;
        foreach ($potem["orderform"]['b'.$b] as $item=>$value) {
            if (($item > 4)&&($b == 1)) {
                $sheet->insertNewRowBefore($row, 1);
//                setMergeCells($sheet,'D'.$row,'E'.$row,,$noborderLeft);
                $conD = 'D'.$row;
                $conE = 'E'.$row;
                $sheet->mergeCells("$conD:$conE");
            }
            $sheet->setCellValue($arr[$a].$row, $value);
//            setCell($sheet,$arr[$a].$row, $value, $styleArray1);
            $row++;
        }
//        $col++;
    }
}

$sheet->getPageSetup()->setFitToPage(true); //将工作表调整为一页

unset($_SESSION['potem3'] ); //注销SESSION

$filenameout = 'PO_'.$potem['shortName'];
outExcel($spreadsheet,$filenameout);


