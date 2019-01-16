<?php
session_start();
header("Content-type: text/html; charset=utf-8");
require_once('autoloadconfig.php');  //判断是否在线
require_once ('img.php');
// modified by fa at 2019.01.15
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;

$intem1 =  $_SESSION['invoiceTem6'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/invoiceTem6.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
//$spreadsheet->getActiveSheet()->getDefaultColumnDimension()->setWidth(20);  //设置默认列宽
$spreadsheet->getActiveSheet()->getDefaultRowDimension()->setRowHeight(50); //行默认高度
$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(9);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(9);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(11);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('H')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('I')->setWidth(15);  //列宽度
$spreadsheet->getDefaultStyle()->getFont()->setSize(10);

$border = \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN;
$h_center = \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER;
$v_center = \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER;
$styleArray1 = [
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
        'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
        'wrapText' => true,
        'ShrinkToFit'=>true,
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
$styleArraycenter = [
    'alignment' => [
        'vertical' => $v_center,
        'horizontal' => $h_center,
    ],
    'borders' => [
        'top' => [
            'borderStyle' => $border,
        ],
        'bottom' => [
            'borderStyle' => $border,
        ],
        'left' => [
            'borderStyle' => $border,
        ],
        'right' => [
            'borderStyle' => $border,
        ],
    ],

];
$sheet = $spreadsheet->getActiveSheet();

for ($i=0;$i<100;$i++){
    $sheet->getRowDimension($i)->setRowHeight(15); //列高度
}

//fill header
$sheet->setCellValue("F5", 'INVOICE NO.'.$intem1['invoicedata']['invoiceNumber']);
$sheet->setCellValue("C6", $intem1['invoicedata']['a1']);
$sheet->setCellValue("C7", $intem1['invoicedata']['a2']);
$sheet->setCellValue("C8", "");
$sheet->setCellValue("C9", 'Attn'.$intem1['invoicedata']['a3']);
$sheet->setCellValue("J7",$intem1['invoicedate']);

//fill bottom
$sheet->setCellValue('D21',$intem1['remark']['bottomremark'][0]);
$sheet->setCellValue('D22',$intem1['remark']['bottomremark'][1]);
$sheet->setCellValue('E29',$intem1['remark']['c1']);
$sheet->setCellValue('E30',$intem1['remark']['c2']);
$sheet->setCellValue('E31',$intem1['remark']['c3']);
$sheet->setCellValue('E32',$intem1['remark']['c4']);
$sheet->setCellValue('E33',$intem1['remark']['c5']);

//fill main content
{
    //form header
    {
        //three description input
        {
            $sheet->setCellValue("D14", $intem1['invoiceform']['ba1'][0]);
            $sheet->setCellValue("D15", $intem1['invoiceform']['ba1'][4]);
            $sheet->setCellValue("D16", $intem1['invoiceform']['ba1'][5]);
        }

        //Unit Price , Ammount , Precent of ammount
        {
            $sheet->setCellValue("I13", $intem1['invoiceform']['ba1'][1]);
            $sheet->setCellValue("J13", $intem1['invoiceform']['ba1'][2]);
            $sheet->setCellValue("K13", $intem1['invoiceform']['ba1'][3]);
        }
    }

    //form footer
    {
        //total pcs and package
        $sheet->setCellValue("B18", $intem1['invoiceform']['coltb']);
        $sheet->setCellValue("B20", $intem1['invoiceform']['ba1'][7]);
        //total ammount
        $sheet->setCellValue("J18", $intem1['invoiceform']['coltc']);
        //total precent of ammount
        $sheet->setCellValue("K18", $intem1['invoiceform']['ba1'][6]);
        //remark
        $sheet->setCellValue("D20", $intem1['invoiceform']['formremark']);
    }

    //form data
    {
        for ($i=$intem1['invoiceform']['brrnum']-1,$j=$intem1['invoiceform']['formnum']-1;$j>=0&&$i>=0;$j--,$i--){
            add_row($intem1['invoiceform'],$i,$j);
        }
    }

}
function add_row($data,$i,$j){
    global $sheet;
    $sheet->insertNewRowBefore(18,1);

    //quantity
    $sheet->setCellValue("A18", "**");
    $sheet->setCellValue("B18", $data['b1'][$j]);
    $sheet->setCellValue("C18", "**PCS");
    //Po No.
    $sheet->setCellValue("D18", "PO No.:  ".$data['b4'][$j]);
    //Color
    $sheet->setCellValue("E18", "COLOUR:  ".$data['b5'][$j]);
    //our job No.
    $sheet->setCellValue("F18", "OUR JOB NO.:  ".$data['b6'][$j]);
    //description
    $sheet->setCellValue("G18", "DESCRIPTION:  ".$data['b7'][$j]);
    //unit price
    $sheet->setCellValue("I18", $data['b8'][$j]);
    //amount
    $sheet->setCellValue("J18", $data['b9'][$j]);
    //precent of amount
    $sheet->setCellValue("K18", $data['b3'][$j]);
}

$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页
unset($_SESSION['invoiceTem6'] ); //注销SESSION
$output=  ($_GET['action'] == 'formdown' )? 1:0;
$nt = date("YmdHis",time()); //转换为日期。
$filenameout = 'intem1out'.$nt.'.xlsx';
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

