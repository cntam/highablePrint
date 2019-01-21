<?php
session_start();
header("Content-type: text/html; charset=utf-8");
//Modified by 俊伟
/*港源行國際有限公司*/

require_once('autoloadconfig.php');  //判断是否在线

require_once ('img.php');

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;

$hanginglistTem1 =  $_SESSION['packinglist'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/hanginglistTem1.xlsx');
$sheet = $spreadsheet->getActiveSheet();
$spreadsheet->getDefaultStyle()->getFont()->setName('Arimo');
//样式，下框细边
$styleArray1 = [
        'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
        'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
        'wrapText' => true,
        'ShrinkToFit'=>true,
    ],
    'borders' => [
        'bottom' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
    ],
];
//填数据
//header
//FORM TO
$sheet->setCellValue('B2', $hanginglistTem1["invoicedata"]["a1"]);
$sheet->setCellValue('B3', $hanginglistTem1["invoicedata"]["a2"]);
$sheet->setCellValue('B4', $hanginglistTem1["invoicedata"]["a3"]);
$sheet->setCellValue('B5', $hanginglistTem1["invoicedata"]["a4"]);
$sheet->setCellValue('B6', $hanginglistTem1["invoicedata"]["a5"]);
$sheet->setCellValue('B7', $hanginglistTem1["invoicedata"]["a6"]);

$sheet->setCellValue('F2', $hanginglistTem1["invoicedata"]["a7"]);
$sheet->setCellValue('F3', $hanginglistTem1["invoicedata"]["a8"]);
$sheet->setCellValue('F4', $hanginglistTem1["invoicedata"]["a9"]);
$sheet->setCellValue('F5', $hanginglistTem1["invoicedata"]["a10"]);
$sheet->setCellValue('F6', $hanginglistTem1["invoicedata"]["a11"]);
$sheet->setCellValue('F7', $hanginglistTem1["invoicedata"]["a12"]);

//PO Number
$sheet->setCellValue('A10', $hanginglistTem1["invoicedata"]["a13"]);
$sheet->setCellValue('C10', $hanginglistTem1["invoicedata"]["a14"]);
$sheet->setCellValue('E10', $hanginglistTem1["invoicedata"]["a15"]);
$sheet->setCellValue('H10', $hanginglistTem1["invoicedata"]["a16"]);

//Single Size Breakdown
$sheet->setCellValue('B17', $hanginglistTem1["invoicedata"]["a17"]);
$sheet->setCellValue('C17', $hanginglistTem1["invoicedata"]["a18"]);
$sheet->setCellValue('D17', $hanginglistTem1["invoicedata"]["a19"]);
$sheet->setCellValue('E17', $hanginglistTem1["invoicedata"]["a20"]);
$sheet->setCellValue('F17', $hanginglistTem1["invoicedata"]["a21"]);
$sheet->setCellValue('G17', $hanginglistTem1["invoicedata"]["a22"]);
$sheet->setCellValue('H17', $hanginglistTem1["invoicedata"]["a23"]);

$sheet->setCellValue('B18', $hanginglistTem1["invoicedata"]["a24"]);
$sheet->setCellValue('C18', $hanginglistTem1["invoicedata"]["a25"]);
$sheet->setCellValue('D18', $hanginglistTem1["invoicedata"]["a26"]);
$sheet->setCellValue('E18', $hanginglistTem1["invoicedata"]["a27"]);
$sheet->setCellValue('F18', $hanginglistTem1["invoicedata"]["a28"]);
$sheet->setCellValue('G18', $hanginglistTem1["invoicedata"]["a29"]);
$sheet->setCellValue('H18', $hanginglistTem1["invoicedata"]["a30"]);
$sheet->setCellValue('I18', $hanginglistTem1["invoicedata"]["a31"]);

//GM
$sheet->setCellValue('J9', 'GM: '.$hanginglistTem1["invoicedata"]["a32"]);
$sheet->setCellValue('J10', 'N.W: '.$hanginglistTem1["invoicedata"]["a33"]);
//size格
$sheet->setCellValue('J11', 'Size: '.$hanginglistTem1["invoiceform"]["b1"][0].'X'.$hanginglistTem1["invoiceform"]["b2"][0].'X'.$hanginglistTem1["invoiceform"]["b3"][0].'CM');
$sheet->setCellValue('J12', 'CBM: '.$hanginglistTem1["invoiceform"]["b4"].'m³');

//动态表格
if ($hanginglistTem1["remark"]["clist"]["cnum"] > 0) {
    for ($a = 1; $a <= 4 ; $a++) {
        $row = 23;
        $col = chr(66 + $a); // C
        foreach ($hanginglistTem1["remark"]["clist"]['c'.$a] as $item => $value) {
            if (($item > 2)&&($a == 1)) {
                $sheet->insertNewRowBefore($row, 1);
            }
            $sheet->setCellValue($col.$row, $value);
            $row++;
        }
    }
}

$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

//unset($_SESSION['potem1'] ); //注销SESSION

$output=  ($_GET['action'] == 'formdown' )? 1:0;
$nt = date("md",time()); //转换为日期。
$filenameout = 'Hanginglist_'.$hanginglistTem1['shortname'].$nt.'.xlsx';
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

