<?php
require_once 'aidenfunc.php';
$potem6 =  $_SESSION['potem6'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/potem6.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$spreadsheet->getActiveSheet()->setTitle("sheet1");
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
//$spreadsheet->getActiveSheet()->getDefaultColumnDimension()->setWidth(20);  //设置默认列宽
$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(25);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(25);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(25);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(20);  //列宽度
//
//$spreadsheet->getActiveSheet()->getColumnDimension('I')->setWidth(15);  //列宽度
$spreadsheet->getDefaultStyle()->getFont()->setSize(7);


//填数据
$spreadsheet->getActiveSheet()->setCellValue('B6', $potem6["tosb"]);
$spreadsheet->getActiveSheet()->setCellValue('F7', $potem6 ["podate"]);
$spreadsheet->getActiveSheet()->setCellValue('F6', $potem6["toaddr"]["a1"]);
$spreadsheet->getActiveSheet()->setCellValue('B7', $potem6["toaddr"]["a2"]);



//中部form
//$spreadsheet->getActiveSheet()->setCellValue('B11', $potem6["toaddr"]["a7"]);
//$spreadsheet->getActiveSheet()->setCellValue('B12', $potem6["orderform"]["midpono"]);

$nowcol = 11;
//$spreadsheet->getActiveSheet()->mergeCells("A{$nowcol}:F{$nowcol}");
//$spreadsheet->getActiveSheet()->setCellValue('A'.$nowcol, '(PO NO:  '.$potem6["orderform"]["midpono"].' 注：請在開發票時把"PO NO"寫上，不可重複)');
////$spreadsheet->getActiveSheet()->setCellValue('I'.$nowcol, $potem6["invoiceform"]["amout"]);
////
//$nowcol++;
//$nowcol++;
//
for($x = 0 ,$c = 1; $x <= $potem6["orderform"]["formnum"]; $x++ ,$c++){

$f19 = 11 + 1 * $x;

$spreadsheet->getActiveSheet()->mergeCells("B{$f19}:D{$f19}");

$formarr = array('A'.$f19,'B'.$f19,'E'.$f19,'F'.$f19,'G'.$f19);

    for($i = 1,$y = 0; $i <= $potem6["orderform"]["brrnum"] ; $i++ ,$y++){

        $sheet->setCellValue($formarr[$y],  $potem6["orderform"]['b'.$i][$x]);

    }


    $nowcol = 11  +   1 * $c;



    if($x >9){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($nowcol, 1);
    }

}



//底部REMARK
$nowcol = $potem6["orderform"]["formnum"] > 9? ($nowcol + 1) : 22;
//$spreadsheet->getActiveSheet()->getCell('A1')->setValue($nowcol);

$spreadsheet->getActiveSheet()->setCellValue('E'.$nowcol, $potem6["toaddr"]["a3"]);
$spreadsheet->getActiveSheet()->setCellValue('F'.$nowcol, $potem6["toaddr"]["a4"]);
$spreadsheet->getActiveSheet()->setCellValue('G'.$nowcol, $potem6["toaddr"]["a5"]);
$nowcol++;

//$spreadsheet->getActiveSheet()->mergeCells("A{$nowcol}:E{$nowcol}");
$spreadsheet->getActiveSheet()->setCellValue('A'.$nowcol, '備註：'.$potem6["remark"]["c1"]);
$nowcol++;
////
////$spreadsheet->getActiveSheet()->mergeCells("A{$nowcol}:E{$nowcol}");
//$spreadsheet->getActiveSheet()->setCellValue('B'.$nowcol, $potem6["remark"]["c1"]);
$nowcol++;
$nowcol++;
//
////
//$spreadsheet->getActiveSheet()->mergeCells("A{$nowcol}:E{$nowcol}");
$spreadsheet->getActiveSheet()->setCellValue('A'.$nowcol, $potem6["remark"]["c2"]);
$nowcol++;
//
////$spreadsheet->getActiveSheet()->mergeCells("A{$nowcol}:E{$nowcol}");
$spreadsheet->getActiveSheet()->setCellValue('A'.$nowcol, 'FAX:'.$potem6["remark"]["c3"]);
$nowcol++;
$spreadsheet->getActiveSheet()->setCellValue('A'.$nowcol, 'E-mail:'.$potem6["remark"]["c4"]);
$nowcol++;
$spreadsheet->getActiveSheet()->setCellValue('A'.$nowcol, '交貨期:'.$potem6["remark"]["c5"]);


$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

unset($_SESSION['potem6'] ); //注销SESSION

$filenameout = 'PO_'.$potem6['shortName'];
outExcel($spreadsheet, $filenameout);

