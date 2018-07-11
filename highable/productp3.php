<?php
session_start();
require_once('autoloadconfig.php');  //判断是否在线

if ($online) {
    require_once '/home/pan/vendor/autoload.php';

} else {
    require_once '/Applications/XAMPP/xamppfiles/htdocs/composer/vendor/autoload.php';
}
require_once('img.php');

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;

$productp3 = $_SESSION['productp3'];


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
        'ShrinkToFit' => true,
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


$sheet->setCellValue('B2', $productp3['doc']);
$sheet->setCellValue('D2', $productp3['styleno']);
$sheet->setCellValue('F2', $productp3['guest']);


$formarr = array('A','B','D','E','F','G');
for($x = 0 ,$c = 1; $c <= count($formarr); $x++ ,$c++){
    $f19 = 4;

    for($i = 1,$y = 0; $i <= $productp3["formnum"] ; $i++ ,$y++){
        $spreadsheet->getActiveSheet()->mergeCells("B{$f19}:C{$f19}");
        $sheet->setCellValue($formarr[$x].$f19,  $productp3["a1"]['c'.$c][$y]);
        $spreadsheet->getActiveSheet()->getStyle($formarr[$x].$f19)->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->getStyle("B{$f19}:C{$f19}")->applyFromArray($styleArray1);  //BC样式
        $f19++;

    }
}



unset($_SESSION['productp3'] ); //注销SESSION


$spreadsheet->getActiveSheet()->getPageSetup()
    ->setOrientation(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE); //打印橫向
$spreadsheet->getActiveSheet()->getPageSetup()
    ->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);//打印橫向 A4
$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

$output = ($_GET['action'] == 'formdown') ? 1 : 0;
$nt = date("YmdHis", time()); //转换为日期。

$filenameout = 'productp3out' . $nt . '.xlsx';
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

    Header("Location:{$MSFILEURL}");
}