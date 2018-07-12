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
use PhpOffice\PhpSpreadsheet\IOFactory; //工廠保存接口
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器

$productp6 =  $_SESSION['productp6'];
//var_dump($productp6);

//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/productp6.xlsx');
$sheet = $spreadsheet->getActiveSheet();

$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft Yahei');
$spreadsheet->getDefaultStyle()->getFont()->setSize(12);
$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(9.5);
$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(9.5);
$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(9.5);
$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(9.5);
$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(9.5);
$spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(9.5);
$spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(9.5);

$sheet->setCellValue('B2',  $productp6['guest']);
$sheet->setCellValue('B3',  $productp6['billdate']);
$sheet->setCellValue('D2',  $productp6['doc']);
$sheet->setCellValue('D3',  $productp6['styleno']);
$sheet->setCellValue('F2',  $productp6['department']);
$sheet->setCellValue('F3',  $productp6['findate']);
$sheet->setCellValue('G3',  $productp6['trans']);


$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($productp6['fab1'])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('A5', $richText);




$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($productp6['fab2'])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('A14', $richText);

$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($productp6['fab3'])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('A22', $richText);


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
/*
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
*/
];
$spreadsheet->getActiveSheet()->getStyle("A5:G11")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle("A14:G19")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle("A22:G36")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

unset($_SESSION['productp6'] ); //注销SESSION

$output=  ($_GET['action'] == 'formdown' )? 1:0;
$nt = date("YmdHis",time()); //转换为日期。
$filenameout = 'productp6out'.$nt.'.xlsx';
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
exit;
