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

$invoiceTem4 =  $_SESSION['invoice'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/invoiceTem4.xlsx');

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

//填数据
$sheet->mergeCells("H21:K21");

$sheet->setCellValue('A6', 'INVOICE NO.'.$invoiceTem4["invoiceno"]);
$sheet->setCellValue('C7', $invoiceTem4["tosb"]);

$row = 8;
$a = 0;
if ($invoiceTem4["invoicedata"]["arrnum"] > 0) {
    for ($x = 1; $x <= $invoiceTem4["invoicedata"]["arrnum"]; $x++) {
        $col = chr(67 + $a); // C
        $sheet->setCellValue($col.$row, $invoiceTem4["invoicedata"]['a'.$x]);
        $row++;
    }
}

$sheet->setCellValue('L11', $invoiceTem4["invoicedate"]);

//中间表格固定内容
$sheet->setCellValue('E14', $invoiceTem4["invoiceform"]["ba1"]["0"]);
$sheet->setCellValue('J14', $invoiceTem4["invoiceform"]["ba1"]["1"]);
$sheet->setCellValue('K14', $invoiceTem4["invoiceform"]["ba1"]["2"]);
$sheet->setCellValue('L14', $invoiceTem4["invoiceform"]["ba1"]["3"]);
$sheet->setCellValue('M15', $invoiceTem4["invoiceform"]["ba1"]["4"]);
$sheet->setCellValue('E20', $invoiceTem4["invoiceform"]["ba1"]["4"]);
$sheet->setCellValue('H21', 'Less'.$invoiceTem4["invoiceform"]["ba1"]["5"].'%DOWN PAYMENT AND CQ COST  BEFORE SHIPMENT');
$sheet->setCellValue('L21', $invoiceTem4["invoiceform"]["ba1"]["6"]);
$sheet->setCellValue('L22', $invoiceTem4["invoiceform"]["ba1"]["7"]);

//底部注释及银行信息
$sheet->setCellValue('E24', $invoiceTem4["invoiceform"]["formremark"]);
$sheet->setCellValue('E27', $invoiceTem4["remark"]["bottomremark"]["0"]);
$sheet->setCellValue('E28', $invoiceTem4["remark"]["bottomremark"]["1"]);
$sheet->setCellValue('G30', $invoiceTem4["remark"]["bottomremark"]["2"]);
$sheet->setCellValue('G31', $invoiceTem4["remark"]["bottomremark"]["3"]);

$sheet->setCellValue('F33', $invoiceTem4["remark"]["c1"]);
$sheet->setCellValue('F34', $invoiceTem4["remark"]["c2"]);
$sheet->setCellValue('F35', $invoiceTem4["remark"]["c3"]);
$sheet->setCellValue('F36', $invoiceTem4["remark"]["c4"]);


//中部表格动态
$row = 17;
foreach ($invoiceTem4["invoiceform"]["b4"] as $item => $value) {
    if ($item > 0) {
        $sheet->insertNewRowBefore($row, 2);
    }
    $sheet->setCellValue('E'.$row, $value);
    $row += 2;
}
//18行BCD列
if ($invoiceTem4["invoiceform"]["brrnum"] > 0) {
    for ($a = 1; $a <= 3 ; $a++) {
        $row = 18;
        $col = chr(65 + $a); // B
        foreach ($invoiceTem4["invoiceform"]['b'.$a] as $value) {
            $sheet->setCellValue($col.$row, $value);
            $row += 2;
        }
    }

}

//18行剩下的列
if ($invoiceTem4["invoiceform"]["brrnum"] > 0) {
    for ($a = 1, $b = 5; $a <= 9 ; $a++, $b++) {
        $row = 18;
        $col = chr(68 + $a); // B
        foreach ($invoiceTem4["invoiceform"]['b'.$b] as $value) {
            $sheet->setCellValue($col.$row, $value);
            $row += 2;
        }
    }

}




$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页
// unset($_SESSION['invoice'] ); //注销SESSION

$output=  ($_GET['action'] == 'formdown' )? 1:0;
$nt = date("YmdHis",time()); //转换为日期。
$filenameout = 'invoiceTem4out'.$nt.'.xlsx';
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

