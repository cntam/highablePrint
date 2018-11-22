<?php
session_start();
require_once('autoloadconfig.php');  //判断是否在线

require_once ('img.php');

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
$sheet->setCellValue('B2',$productp4['guest']);
$sheet->setCellValue('H3',$productp4['styleno']);
$sheet->setCellValue('I4',$productp4['a1']['oarr']['o1']);
$sheet->setCellValue('M4',$productp4['a1']['oarr']['o0']);

$formarr =array('A','B','C','D','E','F','G','H','I','R','S');
for($i = 2;$i<=9;$i++){
    $sheet->setCellValue($formarr[$i-1].'5',$productp4['a1']['oarr']['o'.$i]);
}
for($i = 1;$i<=$productp4['a1']['carr']['colnum'];$i++){
    foreach ($productp4['a1']['carr']['c'.$i] as $index => $item){
        $sheet->setCellValue($formarr[$i-1].(6+$index),$item);
    }
}


$spreadsheet->getActiveSheet()->getPageSetup()
    ->setOrientation(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE); //打印橫向
$spreadsheet->getActiveSheet()->getPageSetup()
    ->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);//打印橫向 A4
$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页
//unset($_SESSION['productp4'] ); //注销SESSION

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

