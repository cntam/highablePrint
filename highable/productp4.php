<?php
session_start();
//require '../vendor/autoload.php';
require '/home/pan/vendor/autoload.php';

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

$sheet->setCellValue('I4',  $productp4['a1']["a1"]);
$sheet->setCellValue('M4',  $productp4['a1']["a2"]);

$formnuma= $productp4["formnum"] +6;
for($i = 6,$a = 0; $i<$formnuma  ;$i++){
    if($formnuma>12 && $i>11 ){
        $y = $i;
        $spreadsheet->getActiveSheet()->insertNewRowBefore($y, 1);

    }

    $sheet->setCellValue("A{$i}", $productp4['a1']["b" . $a][0]);
    $sheet->setCellValue("B{$i}", $productp4['a1']["c". $a][0]);
    $sheet->setCellValue("D{$i}", $productp4['a1']["d". $a][0]);
    $sheet->setCellValue("F{$i}", $productp4['a1']["e". $a][0]);
    $sheet->setCellValue("H{$i}", $productp4['a1']["f". $a][0]);
    $sheet->setCellValue("J{$i}", $productp4['a1']["g". $a][0]);
    $sheet->setCellValue("L{$i}", $productp4['a1']["h". $a][0]);
    $sheet->setCellValue("N{$i}", $productp4['a1']["i". $a][0]);
    $sheet->setCellValue("P{$i}", $productp4['a1']["j". $a][0]);
    $sheet->setCellValue("R{$i}", $productp4['a1']["k". $a][0]);
    $sheet->setCellValue("S{$i}", $productp4['a1']["l". $a][0]);

    for ($v = 0; $v < 19; $v++) {
        $col = chr(97 + $v);
        $spreadsheet->getActiveSheet()->getStyle($col.$i)->applyFromArray($styleArray1);
    }

    $a++;

}
$spreadsheet->getActiveSheet()->getPageSetup()
    ->setOrientation(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE); //打印橫向
$spreadsheet->getActiveSheet()->getPageSetup()
    ->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);//打印橫向 A4
$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页
unset($_SESSION['productp4'] ); //注销SESSION

$output=  ($_GET['action'] == 'formprint' )? 1:0;
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

