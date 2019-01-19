<?php
session_start();
// modified by fa at 2019.01.16

header("Content-type: text/html; charset=utf-8");
require_once('autoloadconfig.php');  //判断是否在线
require_once('img.php');

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;

$intem1 = $_SESSION['invoiceTem12'];

$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/invoiceTem12.xlsx');
$sheet = $spreadsheet->getActiveSheet();
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
$spreadsheet->getActiveSheet()->getColumnDimension('I')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('J')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(22);  //列宽度
$spreadsheet->getDefaultStyle()->getFont()->setSize(10);
$sheet = $spreadsheet->getActiveSheet();


//fill header
{
    $sheet->setCellValue("B7", 'Invoice NO.' . $intem1['invoicedata']['invoiceNumber']);

    $sheet->setCellValue("B8", $intem1['tosb']);//m/s
    $sheet->setCellValue("B9", $intem1['invoicedata']['a1']);
    $sheet->setCellValue("B10", $intem1['invoicedata']['a2']);
    $sheet->setCellValue("B11", $intem1['invoicedata']['a3']);

    $sheet->setCellValue("B13", $intem1['invoicedata']['a4']);//attn
    $sheet->setCellValue("J13", $intem1['invoicedate']);
}
//fill main content
{
    //form header
    {
        //description
        {
            $sheet->setCellValue("F16", $intem1['invoicedata']['a5']);
            $sheet->setCellValue("F17", $intem1['invoicedata']['a6']);
            $sheet->setCellValue("F18", $intem1['invoicedata']['a7']);
            $sheet->setCellValue("F19", $intem1['invoicedata']['a8']);
            $sheet->setCellValue("F20", $intem1['invoicedata']['a9']);
            $sheet->setCellValue("I17", $intem1['invoiceform']['ba1'][0]);
            $sheet->setCellValue("J17", $intem1['invoiceform']['ba1'][1]);
        }

    }

    //form footer
    {
        //total amount
        $sheet->setCellValue("J25", $intem1['invoiceform']['coltc']);
        //remark
        $sheet->setCellValue('F31', $intem1['remark']['bottomremark'][0]);
        $sheet->setCellValue('F33', $intem1['remark']['bottomremark'][1]);
        //bamk info
        $sheet->setCellValue('F36', $intem1['remark']['c1']);
        $sheet->setCellValue('F37', $intem1['remark']['c2']);
        $sheet->setCellValue('F38', $intem1['remark']['c3']);
        $sheet->setCellValue('F39', $intem1['remark']['c4']);

    }

    //form data
    {
        for ($i=$intem1['invoiceform']['brrnum']-1,$j=$intem1['invoiceform']['formnum']-1;$j>=0&&$i>=0;$j--,$i--){
            add_row($intem1['invoiceform'],$i,$j);
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

$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页
//unset($_SESSION['invoiceTem12'] ); //注销SESSION
$output = ($_GET['action'] == 'formdown') ? 1 : 0;
$nt = date("YmdHis", time()); //转换为日期。
$filenameout = 'intem1out' . $nt . '.xlsx';
if ($output) {
    // Redirect output to a client’s web browser (Xlsx)
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment;filename=' . "$filenameout");
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
} else {
    $writer = new Xlsx($spreadsheet);
    $writer->save('../output/' . $filenameout);

    $FILEURL = 'http://allinone321.com/highable/output/' . $filenameout;
    $MSFILEURL = 'http://view.officeapps.live.com/op/view.aspx?src=' . urlencode($FILEURL);
    //echo "<a href= 'http://view.officeapps.live.com/op/view.aspx?src=". urlencode($FILEURL)."' target='_blank' >跳轉--{$filename}</a>";
    Header("Location:{$MSFILEURL}");
};

