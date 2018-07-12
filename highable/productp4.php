<?php
session_start();
require_once('autoloadconfig.php');  //判断是否在线

if($online){
    require_once '/home/pan/vendor/autoload.php';

}else{
    require_once '/Applications/XAMPP/xamppfiles/htdocs/composer/vendor/autoload.php';
}

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;

$productp4 =  $_SESSION['productp4'];
$action = $_GET['action'];

//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/productp4.xlsx');
$sheet = $spreadsheet->getActiveSheet();

$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft Yahei');
$spreadsheet->getDefaultStyle()->getFont()->setSize(10);
$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(15);
for ($v = 1; $v < 19; $v++) {
    $col = chr(97 + $v);
    $spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(9);
}
$styleArray1 = [
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
        'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP,
        'wrapText' => true,
        'ShrinkToFit'=>true,
    ],
    'font' => [
        'Size' => '10',
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

];


$sheet->setCellValue('B2',  $productp4['guest']);
$sheet->setCellValue('B3',  $productp4['billdate']);
$sheet->setCellValue('H2',  $productp4['doc']);
$sheet->setCellValue('H3',  $productp4['styleno']);
$sheet->setCellValue('P2',  $productp4['department']);
$sheet->setCellValue('P3',  $productp4['findate']);
$sheet->setCellValue('R3',  $productp4['trans']);

$sheet->setCellValue('I4',  $productp4['large']['o0']);
$sheet->setCellValue('M4',  $productp4['large']['o1']);

$sheet->setCellValue("B5", $productp4['ct'][13]);
$sheet->setCellValue("D5", $productp4['ct'][14]);
$sheet->setCellValue("F5", $productp4['ct'][15]);
$sheet->setCellValue("H5", $productp4['ct'][16]);
$sheet->setCellValue("J5", $productp4['ct'][17]);
$sheet->setCellValue("L5", $productp4['ct'][18]);
$sheet->setCellValue("N5", $productp4['ct'][19]);
$sheet->setCellValue("P5", $productp4['ct'][20]);



$formarr = array('A','B','D','F','H','J','L','N','P','R','S');
for($x = 0 ,$c = 1; $c <= count($formarr); $x++ ,$c++){
    $f19 = 6;
    for($i = 1,$y = 0; $i <= $productp4["formnumb"] ; $i++ ,$y++){
        $sheet->setCellValue($formarr[$x].$f19,  $productp4['large']['c'.$c][$y]);
        $f19++;
    }
}


for($x = 0 ,$c = 1; $c <= 19; $x++ ,$c++){
    $f19 = 6;
    for($i = 1,$y = 0; $i <= $productp4["formnumb"] ; $i++ ,$y++){
        $col = chr(97 + $x);
        $spreadsheet->getActiveSheet()->getStyle($col.$f19)->applyFromArray($styleArray1);
        $f19++;
    }
}


$spreadsheet->getActiveSheet()->getPageSetup()
    ->setOrientation(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE); //打印橫向
$spreadsheet->getActiveSheet()->getPageSetup()
    ->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);//打印橫向 A4
$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页
unset($_SESSION['productp4'] ); //注销SESSION

$output=  ($_GET['action'] == 'formdown' )? 1:0;
$nt = date("YmdHis",time()); //转换为日期。
$filenameout = 'productp4out'.$nt.'.xlsx';
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

    Header("Location:{$MSFILEURL}");
}

