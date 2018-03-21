<?php
session_start();
//require '../vendor/autoload.php';
require '/home/pan/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;
$productp3 =  $_SESSION['productp3'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/productp3.xlsx');
$sheet = $spreadsheet->getActiveSheet();
$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(15);
$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(16);
$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(52);

$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(16);
$spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(16);
$spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(16);
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


$sheet->setCellValue('B2',  $productp3['doc']);
$sheet->setCellValue('D2',  $productp3['styleno']);
$sheet->setCellValue('F2',  $productp3['guest']);


$formnuma= $productp3["formnum"];
for($i = 0,$a = 0,$row = 4; $i<$formnuma  ;$i++, $row++){
    if($formnuma>25 && $i>24 ){
        $y = $row;
        $spreadsheet->getActiveSheet()->insertNewRowBefore($y, 1);

    }
    $sheet->setCellValue("A{$row}", $productp3['a1']["a" .$a][0]);
    $spreadsheet->getActiveSheet()->mergeCells("B{$row}:C{$row}");
    $sheet->setCellValue("B{$row}", $productp3['a1']["b" .$a][0]);

    $sheet->setCellValue("D{$row}", $productp3['a1']["c". $a][0]);
    //$sheet->setCellValue("D{$row}", $productp3['a1']["d". $a][0]);

    $sheet->setCellValue("E{$row}", $productp3['a1']["d". $a][0]);
    $sheet->setCellValue("F{$row}", $productp3['a1']["e". $a][0]);
    $sheet->setCellValue("G{$row}", $productp3['a1']["f". $a][0]);


    $spreadsheet->getActiveSheet()->getStyle("A{$row}")->applyFromArray($styleArray1);
    //$spreadsheet->getActiveSheet()->getStyle("B{$row}")->applyFromArray($styleArray1);

    $spreadsheet->getActiveSheet()->getStyle("B{$row}:C{$row}")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->getStyle("D{$row}")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->getStyle("E{$row}")->applyFromArray($styleArray1);

    $spreadsheet->getActiveSheet()->getStyle("F{$row}")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->getStyle("G{$row}")->applyFromArray($styleArray1);
    $a++;

}

unset($_SESSION['productp3'] ); //注销SESSION


$spreadsheet->getActiveSheet()->getPageSetup()
    ->setOrientation(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE); //打印橫向
$spreadsheet->getActiveSheet()->getPageSetup()
    ->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);//打印橫向 A4
$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

$output=  ($_GET['action'] == 'formprint' )? 1:0;
$nt = date("YmdHis",time()); //转换为日期。

$filenameout = 'productp3out'.$nt.'.xlsx';
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