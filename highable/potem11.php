<?php
require_once 'aidenfunc.php';

$potem11 =  $_SESSION['potem11'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/potem11.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle("sheet1");
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
//$sheet->getDefaultColumnDimension()->setWidth(20);  //设置默认列宽
//$sheet->getColumnDimension('A')->setWidth(25);  //列宽度
$sheet->getColumnDimension('B')->setWidth(5);  //列宽度
$sheet->getColumnDimension('C')->setWidth(15);  //列宽度
$sheet->getColumnDimension('D')->setWidth(20);  //列宽度
$sheet->getColumnDimension('E')->setWidth(30);  //列宽度
$sheet->getColumnDimension('F')->setWidth(20);  //列宽度
$sheet->getColumnDimension('G')->setWidth(15);  //列宽度
$sheet->getColumnDimension('H')->setWidth(20);  //列宽度
$sheet->getColumnDimension('I')->setWidth(30);  //列宽度
$spreadsheet->getDefaultStyle()->getFont()->setSize(7);

//填数据
$sheet->setCellValue('B11', 'TO: '.$potem11["tosb"]);
$sheet->setCellValue('G11', 'DATE: '.$potem11["podate"]);

$sheet->setCellValue('B16', $potem11["toaddr"]["a1"]);
$sheet->setCellValue('B17', $potem11["toaddr"]["a2"]);
$sheet->setCellValue('B18', $potem11["toaddr"]["a3"]);
$sheet->setCellValue('B19', $potem11["toaddr"]["a4"]);
$sheet->setCellValue('B21', 'Contact (聯絡人) :'.$potem11["toaddr"]["a5"]);
$sheet->setCellValue('B23', 'Tel (電話): '.$potem11["toaddr"]["a6"]);
$sheet->setCellValue('B25', 'Fax (傳真): '.$potem11["toaddr"]["a7"]);
$sheet->setCellValue('B27', 'Email (電郵): '.$potem11["toaddr"]["a8"]);
$sheet->setCellValue('B29', 'Customer PO# (客戶訂單號碼):'.$potem11["toaddr"]["a9"]);
$sheet->setCellValue('B31', 'Payment currency (付款幣值):'.$potem11["toaddr"]["a10"]);


$sheet->setCellValue('G16', $potem11["toaddr"]["a11"]);
$sheet->setCellValue('G17', $potem11["toaddr"]["a12"]);
$sheet->setCellValue('G18', $potem11["toaddr"]["a13"]);
$sheet->setCellValue('G19', $potem11["toaddr"]["a14"]);
$sheet->setCellValue('G21', 'Contact (聯絡人) :'.$potem11["toaddr"]["a15"]);
$sheet->setCellValue('G23', 'Tel (電話): '.$potem11["toaddr"]["a16"]);
$sheet->setCellValue('G25', 'Fax (傳真): '.$potem11["toaddr"]["a17"]);
$sheet->setCellValue('G27', 'Ship mode (運輸方式) :'.$potem11["toaddr"]["a18"]);
$sheet->setCellValue('G29', 'Port of Discharge (港口名稱): '.$potem11["toaddr"]["a19"]);


//中部form

//$nowcol = 35;
////$sheet->mergeCells("A{$nowcol}:F{$nowcol}");
//$sheet->setCellValue('A'.$nowcol, 'PO NO: '.$potem11["orderform"]["midpono"].'   注：請在開發票時把“PONO”寫上，不可重復，并且寫上制單號）');
////$sheet->setCellValue('I'.$nowcol, $potem11["invoiceform"]["amout"]);


for($x = 0 ,$c = 1; $c <= $potem11["orderform"]["formnum"]; $x++ ,$c++){

$f19 = 35 + 1 * $x;

$sheet->mergeCells("B{$f19}:C{$f19}");
$sheet->mergeCells("F{$f19}:G{$f19}");

$formarr = array('B'.$f19,'D'.$f19,'E'.$f19,'F'.$f19,'H'.$f19,'I'.$f19);

    for($i = 1,$y = 0; $i <= $potem11["orderform"]["brrnum"] ; $i++ ,$y++){

        $sheet->setCellValue($formarr[$y],  $potem11["orderform"]['b'.$i][$x]);

    }

    $nowcol = 35  +   1 * $c;


    if($x >1){
        $sheet->insertNewRowBefore($nowcol, 1);
    }

}
$nowcol = $potem11["orderform"]["formnum"] > 1 ? ($nowcol + 2) : 39;
////$sheet->getCell('A1')->setValue($nowcol); 貨送以下地址
////$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
////$sheet->setCellValue('A'.$nowcol, '貨送以下地址');
////$nowcol++;
//
//$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
$sheet->setCellValue('D'.$nowcol, $potem11["toaddr"]["a20"]);
$nowcol++;
$nowcol++;
$nowcol++;

////$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
$sheet->setCellValue('D'.$nowcol, $potem11["toaddr"]["a21"]);
$nowcol++;
$nowcol++;
$nowcol++;
$nowcol++;



$sheet->setCellValue('B'.$nowcol, 'ORDERED BY (經手人) :'.$potem11["remark"]["c1"]);
$sheet->setCellValue('I'.$nowcol, $potem11["remark"]["c2"]);
$nowcol++;
$nowcol++;

$sheet->setCellValue('B'.$nowcol, 'E-MAIL (電郵): '.$potem11["remark"]["c3"]);
$sheet->setCellValue('I'.$nowcol, $potem11["remark"]["c4"]);
////$sheet->mergeCells("A{$nowcol}:E{$nowcol}");
//$sheet->setCellValue('E'.$nowcol, $potem11["remark"]["c3"]);
//$nowcol++;
//
//$sheet->setCellValue('A'.$nowcol, $potem11["remark"]["c4"]);
//
$sheet->getPageSetup()
    ->setOrientation(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_PORTRAIT);  //竖放置
$sheet->getPageSetup()
    ->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);  //A4
$sheet->getPageSetup()->setFitToPage(true); //将工作表调整为一页
//
unset($_SESSION['potem11'] ); //注销SESSION

$filenameout = 'PO_'.$potem11['shortName'].'_'.$potem11['pono'];
outExcel($spreadsheet,$filenameout);
