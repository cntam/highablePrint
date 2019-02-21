<?php
header("Content-type: text/html; charset=utf-8");
/*港源行國際有限公司*/

require_once 'aidenfunc.php';

$potem1 =  $_SESSION['potem1'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/potem1.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle("sheet1");
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
//$sheet->getDefaultColumnDimension()->setWidth(20);  //设置默认列宽
$sheet->getColumnDimension('A')->setWidth(16);  //列宽度
$sheet->getColumnDimension('B')->setWidth(16);  //列宽度
$sheet->getColumnDimension('C')->setWidth(20);  //列宽度
$sheet->getColumnDimension('D')->setWidth(20);  //列宽度
$sheet->getColumnDimension('E')->setWidth(25);  //列宽度
$sheet->getColumnDimension('F')->setWidth(20);  //列宽度
$sheet->getColumnDimension('G')->setWidth(20);  //列宽度
//$sheet->getColumnDimension('H')->setWidth(15);  //列宽度
$sheet->getColumnDimension('I')->setWidth(20);  //列宽度
$spreadsheet->getDefaultStyle()->getFont()->setSize(7);

////填数据
//header
setCell($sheet, "A1", $potem1["remark"]["poheader"]["poheada1"], $noborderCenter);
setCell($sheet, "A2", 'Address:'.$potem1["remark"]["poheader"]["poheada2"], $noborderCenter);
setCell($sheet, "A3", $potem1["remark"]["poheader"]["poheada3"], $noborderCenter);
setCell($sheet,"A4",$potem1["remark"]['poheader']['poheada4'].$potem1["remark"]['poheader']['poheada5'],$noborderCenter);
setCell($sheet, "A5", 'Attn :'.$potem1["remark"]["poheader"]["poheada6"], $noborderCenter);

//$sheet->setCellValue('E5', $potem1["toaddr"]["a8"]);
$sheet->mergeCells("A9:E9");
$sheet->mergeCells("A10:C10");
$sheet->mergeCells("A14:B14");
$sheet->mergeCells("A15:D15");
$sheet->mergeCells("A19:G19");


//$sheet->setCellValue('A5', 'Attn : '.$potem1["toaddr"]["a7"]);
$sheet->setCellValue('A9', 'TO: '.$potem1["toaddr"]["a8"]);
$sheet->setCellValue('I9', $potem1["podate"]);
$sheet->setCellValue('A10', $potem1["tosb"]);
$sheet->setCellValue('A11', $potem1["toaddr"]["a1"]);
$sheet->setCellValue('A13', 'TEL: '.$potem1["toaddr"]["a2"].'  FAX: '.$potem1["toaddr"]["a3"].'  e-mail：'.$potem1["toaddr"]["a4"]);
$sheet->setCellValue('A14', 'ATTN：'.$potem1["toaddr"]["a5"]);

$sheet->setCellValue('A15', 'Email：'.$potem1["toaddr"]["a6"]);
$sheet->setCellValue('I15', $potem1["toaddr"]["a7"]);
$sheet->setCellValue('A19', 'PO NO:'.$potem1["orderform"]["midpono"].'  (注:請在開發票時把“PO NO”寫上,不可重復,并且寫上OUR REF)');
//$sheet->setCellValue('C27', '送货地址：'.$potem1["remark"][c1].PHP_EOL.$potem1["remark"][c2].'收件人'.PHP_EOL.$potem1["remark"][c3]);
$send = '送货地址：'.$potem1["remark"]['c1'].PHP_EOL.$potem1["remark"]['c2'].PHP_EOL.'收件人: '.$potem1["remark"]['c3'];
setCell($sheet,"C27",$send,$noborderCenter);

// Remark
$sheet->setCellValue('B30', $potem1["remark"]["c4"]);

//中间表格

$sheet->setCellValue('I21', 'unit price'.PHP_EOL.'('.$potem1["orderform"]['b7'].')');
$sheet->setCellValue('F21', 'QTY'.$potem1["orderform"]['b6']);
$nowcol = 22;


for($x = 0 ,$c = 1; $x <= $potem1["orderform"]["formnum"]; $x++ ,$c++){

$f19 = 22 + 1 * $x;

$sheet->mergeCells("A{$f19}:B{$f19}");
$sheet->mergeCells("C{$f19}:D{$f19}");
$sheet->mergeCells("F{$f19}:G{$f19}");

$formarr = array('A'.$f19,'C'.$f19,'E'.$f19,'F'.$f19,'I'.$f19);

//    for($i = 1,$y = 0; $i <= $potem1["orderform"]["brrnum"] ; $i++ ,$y++){
    for($i = 1,$y = 0; $i <= count($formarr) ; $i++ ,$y++){
        $sheet->setCellValue($formarr[$y], $potem1["orderform"]['b' . $i][$x]);

    }
    $nowcol = 22  +   1 * $c;

    if($x >3){
        $sheet->insertNewRowBefore($nowcol, 1);
    }

}

$sheet->getPageSetup()
    ->setOrientation(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_PORTRAIT);  //竖放置
$sheet->getPageSetup()
    ->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);  //A4
$sheet->getPageSetup()->setFitToPage(true); //将工作表调整为一页

unset($_SESSION['potem1'] ); //注销SESSION

//$filenameout = 'PO_'.$potem1['shortName'].'_'.$potem1['pono'];
$filenameout = 'PO_'.$potem1['pono'];
outExcel($spreadsheet,$filenameout);

