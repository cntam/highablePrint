<?php
require_once 'aidenfunc.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;

$invoice =  $_SESSION['invoice'];


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
$sheet->setCellValue("A1",$invoice['remark']['poheader']['poheada1']);
setCell($sheet,"A2",$invoice["remark"]['poheader']['poheada2'],$noborderCenter);
setCell($sheet,"A3",$invoice["remark"]['poheader']['poheada3'],$noborderCenter);
setCell($sheet,"A4",'Tel:'.$invoice["remark"]['poheader']['poheada5'].'Fax:'.$invoice["remark"]['poheader']['poheada5'],$noborderCenter);


$sheet->mergeCells("H21:K21");

$sheet->setCellValue('A6', 'INVOICE NO.'.$invoice["invoiceno"]);
$sheet->setCellValue('C7', $invoice["tosb"]);

$row = 8;
$a = 0;
if ($invoice["invoicedata"]["arrnum"] > 0) {
    for ($x = 1; $x <= $invoice["invoicedata"]["arrnum"]; $x++) {
        $col = chr(67 + $a); // C
        $sheet->setCellValue($col.$row, $invoice["invoicedata"]['a'.$x]);
        $row++;
    }
}

$sheet->setCellValue('L11', $invoice["invoicedate"]);

//中间表格固定内容
$sheet->setCellValue('E14', $invoice["invoiceform"]["ba1"]["0"]);
$sheet->setCellValue('J14', $invoice["invoiceform"]['b16']);
$sheet->setCellValue('K14', $invoice["invoiceform"]['b17']);
$sheet->setCellValue('L14', $invoice["invoiceform"]['b18']);

$m15 = $invoice["invoiceform"]['b20'][0] / 100 ;
$sheet->setCellValue('M15', $m15);
$sheet->setCellValue('E20', $invoice["invoiceform"]['b19'][0]);
$sheet->setCellValue('H21', 'Less '.$invoice["invoiceform"]['b20'][0].'%DOWN PAYMENT AND CQ COST  BEFORE SHIPMENT');
$sheet->setCellValue('L21', $invoice["invoiceform"]['b21'][0]);
$sheet->setCellValue('L22', $invoice["invoiceform"]['b22'][0]);

//底部注释及银行信息
$sheet->setCellValue('E24', $invoice["invoiceform"]["formremark"]);
$sheet->setCellValue('E27', $invoice["remark"]["bottomremark"]["0"]);
$sheet->setCellValue('E28', $invoice["remark"]["bottomremark"]["1"]);
$sheet->setCellValue('G30', $invoice["remark"]["bottomremark"]["2"]);
$sheet->setCellValue('G31', $invoice["remark"]["bottomremark"]["3"]);


$sheet->setCellValue('F33', $invoice["remark"]["c1"]);
$sheet->setCellValue('F34', $invoice["remark"]["c2"]);
$sheet->setCellValue('F35', $invoice["remark"]["c3"]);
$sheet->setCellValue('F36', $invoice["remark"]["c4"]);
$sheet->setCellValue('G38', $invoice["remark"]["c5"]);


//中部表格动态
$row = 17;
foreach ($invoice["invoiceform"]["b4"] as $item => $value) {
    if ($item > 0) {
        $sheet->insertNewRowBefore($row, 2);
    }
    $sheet->setCellValue('E'.$row, $value);
    $row += 2;
}
//18行BCD列
if ($invoice["invoiceform"]["brrnum"] > 0) {
    for ($a = 1; $a <= 3 ; $a++) {
        $row = 18;
        $col = chr(65 + $a); // B
        foreach ($invoice["invoiceform"]['b'.$a] as $value) {
            $sheet->setCellValue($col.$row, $value);
            $row += 2;
        }
    }

}

//18行剩下的列
if ($invoice["invoiceform"]["brrnum"] > 0) {
    for ($a = 1, $b = 5; $a <= 9 ; $a++, $b++) {
        $row = 18;
        $col = chr(68 + $a); // B
        foreach ($invoice["invoiceform"]['b'.$b] as $value) {
            $sheet->setCellValue($col.$row, $value);
            $row += 2;
        }
    }

}


//$spreadsheet->getActiveSheet()->getPageSetup()
//    ->setOrientation(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE); //打印橫向
$spreadsheet->getActiveSheet()->getPageSetup()
    ->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);//打印橫向 A4
$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

//unset($_SESSION['invoice'] ); //注销SESSION

$filenameout = "Invoice_".$invoice['shortname'];
outExcel($spreadsheet,$filenameout);