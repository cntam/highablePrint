<?php
require_once 'aidenfunc.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;

$inv = $_SESSION['invoice'];

$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/invoiceTem12.xlsx');
$sheet = $spreadsheet->getActiveSheet();
$spreadsheet->getActiveSheet()->setTitle("sheet1");
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
$spreadsheet->getActiveSheet()->getColumnDimension('I')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('J')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(45);  //列宽度
$spreadsheet->getDefaultStyle()->getFont()->setSize(10);
$sheet = $spreadsheet->getActiveSheet();

$sheet->setCellValue("B1",$inv['remark']['poheader']['poheada1']);
setCell($sheet,'B2',$inv['remark']['poheader']['poheada2'],$noborderCenter);
setCell($sheet,'B3',$inv['remark']['poheader']['poheada3'],$noborderCenter);
setCell($sheet,'B4',$inv['remark']['poheader']['poheada4'],$noborderCenter);
$tel = $inv['remark']['poheader']['poheada5'].'  '.$inv['remark']['poheader']['poheada6'];
setCell($sheet,'B5',$tel,$noborderCenter);
//setCell($sheet,'A6',$inv['remark']['poheader']['poheada5'],$noborderCenter);

//fill header
{
    $sheet->setCellValue("B7", 'Invoice NO.' . $inv['invoicedata']['invoiceNumber']);

    $sheet->setCellValue("B8", $inv['tosb']);//m/s
    $sheet->setCellValue("B9", $inv['invoicedata']['a1']);
    $sheet->setCellValue("B10", $inv['invoicedata']['a2']);
    $sheet->setCellValue("B11", $inv['invoicedata']['a3']);

    $sheet->setCellValue("B13", $inv['invoicedata']['a4']);//attn
    $sheet->setCellValue("J13", $inv['invoicedate']);
}
//fill main content
{
    //form header
    {
        //description
        {
            $sheet->setCellValue("F16", $inv['invoicedata']['a5']);
            $sheet->setCellValue("F17", $inv['invoicedata']['a6']);
            $sheet->setCellValue("F18", $inv['invoicedata']['a7']);
            $sheet->setCellValue("F19", $inv['invoicedata']['a8']);
            $sheet->setCellValue("F20", $inv['invoicedata']['a9']);
            $sheet->setCellValue("I17", $inv['invoiceform']['ba1'][0]);
            $sheet->setCellValue("J17", $inv['invoiceform']['ba1'][1]);

            setCell($sheet,'E28',$inv['invoiceform']['formremark'],$noborderLeft);
        }

    }

    //form footer
    {
        //total amount
        $sheet->setCellValue("J25", $inv['invoiceform']['coltc']);
        //remark
        $sheet->setCellValue('F31', $inv['remark']['bottomremark'][0]);
        $sheet->setCellValue('F33', $inv['remark']['bottomremark'][1]);
        //bamk info
        $sheet->setCellValue('F36', $inv['remark']['c1']);
        $sheet->setCellValue('F37', $inv['remark']['c2']);
        $sheet->setCellValue('F38', $inv['remark']['c3']);
        $sheet->setCellValue('F39', $inv['remark']['c4']);

    }

    //form data
    {
        for ($i=$inv['invoiceform']['brrnum']-1,$j=$inv['invoiceform']['formnum']-1;$j>=0&&$i>=0;$j--,$i--){
            add_row($inv['invoiceform'],$i,$j);
        }
    }

}

function add_row($data,$i,$j)
{
    global $sheet;
    $sheet->insertNewRowBefore(24, 1);

    //quantity
    $sheet->setCellValue("A24", $data['b1'][$j]);
    $sheet->setCellValue("B24", 'CTINS');
    $sheet->setCellValue("C24", $data['b3'][$j]);
    $sheet->setCellValue("D24", 'MTSP');
    //description
    $sheet->setCellValue("E24", $data['b5'][$j]);
    $sheet->setCellValue("F24", $data['b6'][$j]);
    $sheet->setCellValue("G24", $data['b7'][$j]);
    $sheet->setCellValue("H24", $data['b8'][$j]);
    //unit price amount
    $sheet->setCellValue("I24", $data['b9'][$j]);
    $sheet->setCellValue("J24", $data['b10'][$j]);
}

$spreadsheet->getActiveSheet()->getPageSetup()
    ->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);  //A4
$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页
unset($_SESSION['invoice'] ); //注销SESSION

$filenameout = "Invoice_".$inv['shortname'];
outExcel($spreadsheet,$filenameout);

