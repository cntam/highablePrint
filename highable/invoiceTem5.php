<!--Template Name: invoiceTem5  -->
<!--PS-->
<!--Modified by 俊伟-->
<!--(Updated by Lau at 2018-11-21)-->
<?php
session_start();
header("Content-type: text/html; charset=utf-8");


require_once('autoloadconfig.php');  //判断是否在线

require_once ('img.php');

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;

$invoiceTem5 =  $_SESSION['invoice'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/invoiceTem5.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle("sheet1");
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
//$spreadsheet->getActiveSheet()->getDefaultColumnDimension()->setWidth(20);  //设置默认列宽
//$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(16);  //列宽度
//$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(16);  //列宽度
//$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(16);  //列宽度
//$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(16);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(40);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('H')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('I')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('J')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('K')->setWidth(20);  //列宽度
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

//FILL SHEET HEADER
{
    $sheet->setCellValue('A6', 'INVOICE NO.' . $invoiceTem5["invoiceno"]);

    $sheet->setCellValue('C7', $invoiceTem5["tosb"]);
    $sheet->setCellValue('C8', $invoiceTem5["invoicedata"]["a1"]);
    $sheet->setCellValue('C9', $invoiceTem5["invoicedata"]["a2"]);

    $sheet->setCellValue('L9', $invoiceTem5["invoicedate"]);

    $sheet->setCellValue('G11', $invoiceTem5["tosb"]);
    $sheet->setCellValue('G12', $invoiceTem5["invoicedata"]["a3"]);
    $sheet->setCellValue('G13', $invoiceTem5["invoicedata"]["a4"]);
    $sheet->setCellValue('G14', $invoiceTem5["invoicedata"]["a5"]);
    $sheet->setCellValue('G15', $invoiceTem5["invoicedata"]["a6"]);
    $sheet->setCellValue('G16', $invoiceTem5["invoicedata"]["a7"]);

    $sheet->setCellValue('K16', $invoiceTem5["invoiceform"]["ba1"][0]);
    $sheet->setCellValue('L16', $invoiceTem5["invoiceform"]["ba1"][1]);
}

//$sheet->setCellValue('K16', $invoiceTem5["invoicedata"]["a6"]);
//$sheet->setCellValue('L16', $invoiceTem5["invoicedata"]["a6"]);
$sheet->setCellValue('M16', $invoiceTem5["invoicedata"]["a8"].'%');

////中间表格固定内容
$sheet->setCellValue('A40', $invoiceTem5["invoicedata"]["a9"]);
$sheet->setCellValue('B40', $invoiceTem5["invoicedata"]["a10"]);
$sheet->setCellValue('C40', $invoiceTem5["invoicedata"]["a11"]);
$sheet->setCellValue('D40', $invoiceTem5["invoicedata"]["a12"]);

//底部注释及银行信息
$sheet->setCellValue('G41', 'Less'.$invoiceTem5["invoicedata"]["a8"].'%DOWN PAYMENT AND CQ COST  BEFORE SHIPMENT');

$sheet->setCellValue('L41', $invoiceTem5["invoicedata"]["a14"]);
$sheet->setCellValue('L42', $invoiceTem5["invoicedata"]["a15"]);

$sheet->setCellValue('E48', 'ORIGIN OF ORIGIN:'.$invoiceTem5["remark"]["bottomremark"]["0"]);
$sheet->setCellValue('G50', $invoiceTem5["remark"]["bottomremark"]["1"]);

$sheet->setCellValue('F53', $invoiceTem5["remark"]["c1"]);
$sheet->setCellValue('F54', $invoiceTem5["remark"]["c2"]);
$sheet->setCellValue('F55', $invoiceTem5["remark"]["c3"]);
$sheet->setCellValue('F56', $invoiceTem5["remark"]["c4"]);

////中部表格动态
$row = 20;
foreach ($invoiceTem5["invoiceform"]["b1"] as $item => $value) {
    if ($item > 4) {
        $sheet->insertNewRowBefore($row , 3);
    }
    $sheet->setCellValue('E'.$row, $value);
    $row += 4;
}
//21行
if ($invoiceTem5["invoiceform"]["formnum"] > 0) {
    for ($a = 1; $a <= 13 ; $a++) {
        $row = 21;
        $col = chr(64 + $a); // A
        foreach ($invoiceTem5["invoiceform"]['b'.($a + 1)] as $value) {
            $sheet->setCellValue($col.$row, $value);
            $row += 4;
        }
    }

}
//
//22行
if (count($invoiceTem5["invoiceform"]["formnum"]) > 0) {
    for ($a = 1, $b = 15; $a <= 11  ; $a++, $b++) {
        $row = 22;
        $col = chr(66 + $a); // C
        foreach ($invoiceTem5["invoiceform"]['b'.$b] as $value) {
            $sheet->setCellValue($col.$row, $value);
            $row += 4;
        }
    }

}




//$sheet->setCellValue('E21', stripcslashes($invoiceTem5["invoiceform"]["b1"]["0"]));





$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页
// unset($_SESSION['invoice'] ); //注销SESSION

$output=  ($_GET['action'] == 'formdown' )? 1:0;
$nt = date("YmdHis",time()); //转换为日期。
$filenameout = 'invoiceTem5out'.$nt.'.xlsx';
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

