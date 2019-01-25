<?php
header("Content-type: text/html; charset=utf-8");
/*港源行國際有限公司*/

require_once 'aidenfunc.php';

$potem1 =  $_SESSION['potem1'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/potem1.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$spreadsheet->getActiveSheet()->setTitle("sheet1");
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
//$spreadsheet->getActiveSheet()->getDefaultColumnDimension()->setWidth(20);  //设置默认列宽
$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(16);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(16);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(25);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(20);  //列宽度
//$spreadsheet->getActiveSheet()->getColumnDimension('H')->setWidth(15);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('I')->setWidth(20);  //列宽度
$spreadsheet->getDefaultStyle()->getFont()->setSize(7);



setCell($sheet,"A1",$potem1["remark"]['poheader']['poheada1'],$noborderCenter);
setCell($sheet,"A2",$potem1["remark"]['poheader']['poheada2'],$noborderCenter);
setCell($sheet,"A3",$potem1["remark"]['poheader']['poheada3'],$noborderCenter);
setCell($sheet,"A4",'Tel:'.$potem1["remark"]['poheader']['poheada5'].'Fax:'.$potem1["remark"]['poheader']['poheada5'],$noborderCenter);
setCell($sheet,"A5",'Attn :'.$potem1["toaddr"]["a9"],$noborderCenter);

////填数据
//$spreadsheet->getActiveSheet()->setCellValue('E5', $potem1["toaddr"]["a8"]);
$spreadsheet->getActiveSheet()->mergeCells("A9:E9");
$spreadsheet->getActiveSheet()->mergeCells("A10:C10");
$spreadsheet->getActiveSheet()->mergeCells("A14:B14");
$spreadsheet->getActiveSheet()->mergeCells("A15:D15");
$spreadsheet->getActiveSheet()->mergeCells("A19:G19");


//$spreadsheet->getActiveSheet()->setCellValue('A5', 'Attn : '.$potem1["toaddr"]["a7"]);
$spreadsheet->getActiveSheet()->setCellValue('A9', 'TO: '.$potem1["toaddr"]["a8"]);
$spreadsheet->getActiveSheet()->setCellValue('I9', $potem1["podate"]);
$spreadsheet->getActiveSheet()->setCellValue('A10', $potem1["tosb"]);
$spreadsheet->getActiveSheet()->setCellValue('A11', $potem1["toaddr"]["a1"]);
$spreadsheet->getActiveSheet()->setCellValue('A13', 'TEL: '.$potem1["toaddr"]["a2"].'  FAX: '.$potem1["toaddr"]["a3"].'  e-mail：'.$potem1["toaddr"]["a4"]);
$spreadsheet->getActiveSheet()->setCellValue('A14', 'ATTN：'.$potem1["toaddr"]["a5"]);

$spreadsheet->getActiveSheet()->setCellValue('A15', 'Email：'.$potem1["toaddr"]["a6"]);
$spreadsheet->getActiveSheet()->setCellValue('I15', $potem1["toaddr"]["a7"]);
$spreadsheet->getActiveSheet()->setCellValue('A19', 'PO NO:'.$potem1["orderform"]["midpono"].'  (注:請在開發票時把“PO NO”寫上,不可重復,并且寫上OUR REF)');
//$spreadsheet->getActiveSheet()->setCellValue('C27', '送货地址：'.$potem1["remark"][c1].PHP_EOL.$potem1["remark"][c2].'收件人'.PHP_EOL.$potem1["remark"][c3]);
$send = '送货地址：'.$potem1["remark"]['c1'].PHP_EOL.$potem1["remark"]['c2'].PHP_EOL.'收件人: '.$potem1["remark"]['c3'];
setCell($sheet,"C27",$send,$noborderCenter);

// Remark
$spreadsheet->getActiveSheet()->setCellValue('B30', $potem1["remark"]["c4"]);

//中间表格

$spreadsheet->getActiveSheet()->setCellValue('I21', 'unit price'.PHP_EOL.'('.$potem1["orderform"]['b7'].')');
$spreadsheet->getActiveSheet()->setCellValue('F21', 'QTY'.$potem1["orderform"]['b6']);
$nowcol = 22;


for($x = 0 ,$c = 1; $x <= $potem1["orderform"]["formnum"]; $x++ ,$c++){

$f19 = 22 + 1 * $x;

$spreadsheet->getActiveSheet()->mergeCells("A{$f19}:B{$f19}");
$spreadsheet->getActiveSheet()->mergeCells("C{$f19}:D{$f19}");
$spreadsheet->getActiveSheet()->mergeCells("F{$f19}:G{$f19}");

$formarr = array('A'.$f19,'C'.$f19,'E'.$f19,'F'.$f19,'I'.$f19);

//    for($i = 1,$y = 0; $i <= $potem1["orderform"]["brrnum"] ; $i++ ,$y++){
    for($i = 1,$y = 0; $i <= count($formarr) ; $i++ ,$y++){
        $sheet->setCellValue($formarr[$y], $potem1["orderform"]['b' . $i][$x]);

    }
    $nowcol = 22  +   1 * $c;

    if($x >3){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($nowcol, 1);
    }

}

$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

//unset($_SESSION['potem1'] ); //注销SESSION

$filenameout = 'PO_'.$potem1['shortName'];
outExcel($spreadsheet,$filenameout);

