<?php
session_start();
header("Content-type: text/html; charset=utf-8");
//KM  && NEXT

require_once('autoloadconfig.php');  //判断是否在线

require_once ('img.php');

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;

$invoiceTem9 =  $_SESSION['invoice'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/invoiceTem9.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$spreadsheet->getActiveSheet()->setTitle("sheet1");
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
//$spreadsheet->getActiveSheet()->getDefaultColumnDimension()->setWidth(20);  //设置默认列宽
//$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(16);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(16);  //列宽度
//$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(16);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(30);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(16);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(16);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(16);  //列宽度
//$spreadsheet->getActiveSheet()->getColumnDimension('K')->setWidth(16);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('I')->setWidth(18);  //列宽度
//$spreadsheet->getActiveSheet()->getColumnDimension('J')->setWidth(16);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('L')->setWidth(20);  //列宽度
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
//$sheet->mergeCells("J1:K1");
//
$sheet->setCellValue('A8', $invoiceTem9["invoicedata"]["a13"]);
$spreadsheet->getActiveSheet()->setCellValue('E8', $invoiceTem9["invoicedata"]["a14"]);
$spreadsheet->getActiveSheet()->setCellValue('I7', 'Invoice No.'.$invoiceTem9["invoicedata"]["a10"]);

$spreadsheet->getActiveSheet()->setCellValue('J11', $invoiceTem9["invoicedate"]);

$spreadsheet->getActiveSheet()->setCellValue('A18', $invoiceTem9["invoicedata"]["a2"]);
$spreadsheet->getActiveSheet()->setCellValue('C18', $invoiceTem9["invoicedata"]["a3"]);
$spreadsheet->getActiveSheet()->setCellValue('E18', $invoiceTem9["invoicedata"]["a4"]);
$spreadsheet->getActiveSheet()->setCellValue('G18', $invoiceTem9["invoicedata"]["a5"]);
$spreadsheet->getActiveSheet()->setCellValue('I18', $invoiceTem9["invoicedata"]["a6"]);
$spreadsheet->getActiveSheet()->setCellValue('K18', $invoiceTem9["invoicedata"]["a7"]);

$spreadsheet->getActiveSheet()->setCellValue('A21', $invoiceTem9["invoicedata"]["a8"]);
$spreadsheet->getActiveSheet()->setCellValue('C21', $invoiceTem9["invoicedata"]["a9"]);
$spreadsheet->getActiveSheet()->setCellValue('E21', $invoiceTem9["invoicedata"]["a10"]);
$spreadsheet->getActiveSheet()->setCellValue('G21', $invoiceTem9["invoicedata"]["a11"]);
$spreadsheet->getActiveSheet()->setCellValue('I21', $invoiceTem9["invoicedata"]["a12"]);


// 中间表格
$spreadsheet->getActiveSheet()->setCellValue('J24', '('.$invoiceTem9["invoiceform"]["b12"].')');
$spreadsheet->getActiveSheet()->setCellValue('L24', '('.$invoiceTem9["invoiceform"]["b13"].')');

$spreadsheet->getActiveSheet()->setCellValue('L32', $invoiceTem9["invoiceform"]["coltb"]);
$spreadsheet->getActiveSheet()->setCellValue('C35', $invoiceTem9["invoiceform"]["formremark"]);

// BOTTOM
$spreadsheet->getActiveSheet()->setCellValue('D40', $invoiceTem9["remark"]["bottomremark"]["0"]);
$spreadsheet->getActiveSheet()->setCellValue('D42', $invoiceTem9["remark"]["bottomremark"]["1"]);

//动态部分
$nowcol = 27;

foreach ($invoiceTem9["invoiceform"]["b1"] as $item => $value) {
    if ($item > 0) {
        $spreadsheet->getActiveSheet()->insertNewRowBefore($nowcol, 4);
    }
    $spreadsheet->getActiveSheet()->setCellValue('A'.$nowcol, $value);
    $nowcol += 4;
}

//$nowcol = 28;
////foreach ($invoiceTem9["invoiceform"] as $item => $value) {
////    $formarr = array('A'.$nowcol,'B'.$nowcol,'D'.$nowcol,'E'.$nowcol,'I'.$nowcol,'J'.$nowcol,'L'.$nowcol);
////    for ($i = 2, $y = 0; $i <= 8; $i++, $y++) {
////        $spreadsheet->getActiveSheet()->setCellValue($formarr[$y], $value['b'.$i]);
////    }
////    $nowcol += 4;
////}
//
$nowcol = 28;
foreach ($invoiceTem9["invoiceform"]["b2"] as $item => $value) {

    $spreadsheet->getActiveSheet()->setCellValue('A'.$nowcol, $value);
    $nowcol += 4;
}

$nowcol = 28;
foreach ($invoiceTem9["invoiceform"]["b3"] as $item => $value) {

    $spreadsheet->getActiveSheet()->setCellValue('B'.$nowcol, $value);
    $nowcol += 4;
}
$nowcol = 28;
foreach ($invoiceTem9["invoiceform"]["b4"] as $item => $value) {

    $spreadsheet->getActiveSheet()->setCellValue('D'.$nowcol, $value);
    $nowcol += 4;
}

$nowcol = 28;
foreach ($invoiceTem9["invoiceform"]["b5"] as $item => $value) {

    $spreadsheet->getActiveSheet()->setCellValue('E'.$nowcol, $value);
    $nowcol += 4;
}
$nowcol = 28;
foreach ($invoiceTem9["invoiceform"]["b6"] as $item => $value) {

    $spreadsheet->getActiveSheet()->setCellValue('I'.$nowcol, $value);
    $nowcol += 4;
}
$nowcol = 28;
foreach ($invoiceTem9["invoiceform"]["b7"] as $item => $value) {

    $spreadsheet->getActiveSheet()->setCellValue('J'.$nowcol, $value);
    $nowcol += 4;
}
$nowcol = 28;
foreach ($invoiceTem9["invoiceform"]["b8"] as $item => $value) {
    $spreadsheet->getActiveSheet()->setCellValue('L'.$nowcol, $value);
    $nowcol += 4;
}


$nowcol = 29;
foreach ($invoiceTem9["invoiceform"]["b9"] as $item => $value) {
    $spreadsheet->getActiveSheet()->setCellValue('A'.$nowcol, $value);
    $nowcol += 4;
}

$nowcol = 30;
foreach ($invoiceTem9["invoiceform"]["b10"] as $item => $value) {
//    if ($item > 1) {
//        $spreadsheet->getActiveSheet()->insertNewRowBefore($nowcol, 4);
//    }
    $spreadsheet->getActiveSheet()->setCellValue('E'.$nowcol, $value);
    $nowcol += 4;
}


$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

// unset($_SESSION['invoice'] ); //注销SESSION

$output=  ($_GET['action'] == 'formdown' )? 1:0;
$nt = date("md",time()); //转换为日期。
$filenameout = "Invoice_{$invoiceTem9['shortname']}_".$nt.'.xlsx';
if($output){
    // Redirect output to a client’s web browser (Xlsx)
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment;filename='."$filenameout");
    header('Cache-Control: max-age=0');
// If you're serving to IE 9, then the following may be needed
    header('Cache-Control: max-age=1');

// If you're serving to IE over SSL, then the following may be needed
    header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
    header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
    header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
    header('Pragma: public'); // HTTP/1.0

    $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
    $writer->save('php://output');
}else{
    $writer = new Xlsx($spreadsheet);
    $writer->save('../output/'.$filenameout);

    $FILEURL = 'http://allinone321.com/highable/output/'.$filenameout;
    $MSFILEURL = 'http://view.officeapps.live.com/op/view.aspx?src='. urlencode($FILEURL);
    //echo "<a href= 'http://view.officeapps.live.com/op/view.aspx?src=". urlencode($FILEURL)."' target='_blank' >跳轉--{$filename}</a>";
    Header("Location:{$MSFILEURL}");
};

