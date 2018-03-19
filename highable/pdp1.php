<?php
session_start();
header("Content-type: text/html; charset=utf-8");
//require '../vendor/autoload.php';
require '/home/pan/vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;

$pdp1 =  $_SESSION['pdp1'];
//var_dump($pdp1);

//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/pdp1.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
$spreadsheet->getDefaultStyle()->getFont()->setSize(10);

$sheet->setCellValue('A3',  $pdp1["SPL_1_code"]);
$sheet->setCellValue('B3',  $pdp1["SPL_1_name"]);
$sheet->setCellValue('C3',  $pdp1["SPL_1_country"]);
$sheet->setCellValue('D3',  $pdp1["SPL_1_contact"]);
$sheet->setCellValue('E3',  $pdp1["SPL_1_address"]);
$sheet->setCellValue('F3',  $pdp1["SPL_1_email"].'/'.$pdp1["SPL_1_tel"].'/'.$pdp1["SPL_1_mobile"].'/'.$pdp1["SPL_1_qq"]);
$sheet->setCellValue('G3',  $pdp1["SPL_1_goods"]);

$sheet->setCellValue('A4',  $pdp1["SPL_2_code"]);
$sheet->setCellValue('B4',  $pdp1["SPL_2_name"]);
$sheet->setCellValue('C4',  $pdp1["SPL_2_country"]);
$sheet->setCellValue('D4',  $pdp1["SPL_2_contact"]);
$sheet->setCellValue('E4',  $pdp1["SPL_2_address"]);
$sheet->setCellValue('F4',  $pdp1["SPL_2_email"].'/'.$pdp1["SPL_2_tel"].'/'.$pdp1["SPL_2_mobile"].'/'.$pdp1["SPL_2_qq"]);
$sheet->setCellValue('G4',  $pdp1["SPL_2_goods"]);

$sheet->setCellValue('F8',  $pdp1["FR_date"]);
$sheet->setCellValue('F10',  $pdp1["FR_ihkno"]);
$sheet->setCellValue('F12',  $pdp1["FR_supplier"]);
$sheet->setCellValue('F14',  $pdp1["FR_suppliercode"]);
$sheet->setCellValue('F16',  $pdp1["FR_comp"]);
$sheet->setCellValue('F18',  $pdp1["FR_width"]);
$sheet->setCellValue('F20',  $pdp1["FR_weight"]);
$sheet->setCellValue('F22',  $pdp1["FR_remark"]);

$sheet->setCellValue('F26',  $pdp1["SO_date"]);
$sheet->setCellValue('F28',  $pdp1["SO_category"]);
$sheet->setCellValue('F30',  $pdp1["SO_styleno"]);
$sheet->setCellValue('F32',  $pdp1["SO_client"]);
$sheet->setCellValue('F34',  $pdp1["SO_fabric"]);
$sheet->setCellValue('F36',  $pdp1["SO_fabricinfo"]);
$sheet->setCellValue('F38',  $pdp1["SO_lining"]);
$sheet->setCellValue('F40',  $pdp1["SO_lininginfo"]);
$sheet->setCellValue('F42',  $pdp1["SO_trim"]);
$sheet->setCellValue('F44',  $pdp1["SO_triminfo"]);
$sheet->setCellValue('F46',  $pdp1["SO_remark"]);


$img = $pdp1["FR_img"];
$img = imagecreatefromjpeg($img);
$width = imagesx($img);
$height = imagesy($img);


// Generate an image
//$gdImage = @imagecreatetruecolor($width, $height) or die('Cannot Initialize new GD image stream');
//$textColor = imagecolorallocate($gdImage, 255, 255, 255);
//imagestring($gdImage, 1, 5, 5,  'Created with PhpSpreadsheet', $textColor);

// Add a drawing to the worksheet
$drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
$drawing->setName('FABRIC RECODE');
$drawing->setDescription('FABRIC RECODE');
//$drawing->setImageResource($gdImage);
$drawing->setImageResource($img);
$drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
$drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
//$drawing->setHeight($width);

//$drawing->setHeight($width>550 ? 550:$width);
$drawing->setWidth(250);
//$drawing->setHeight(150);
$drawing->setCoordinates('A8');
$drawing->setOffsetX(5);
$drawing->setOffsetY(5);
$drawing->setWorksheet($spreadsheet->getActiveSheet());

/*$spreadsheet->getActiveSheet()
    ->getColumnDimension('A')
    ->setWidth(48);
$spreadsheet->getActiveSheet()
    ->getRowDimension(1)
    ->setRowHeight(-1);*/
/*
$spreadsheet->getActiveSheet()->getStyle("A".$listrow)
    ->getAlignment()
    ->setWrapText(true);
$spreadsheet->getActiveSheet()->getStyle("A".$listrow)
    ->getAlignment()
    ->setShrinkToFit(true);
*/
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
$spreadsheet->getActiveSheet()->getStyle("A3:G3")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle("A4:G4")->applyFromArray($styleArray1);

//$spreadsheet->getActiveSheet()->getStyle("A".$listrow)->getFont()->setSize(8);



$img = $pdp1["SO_img"];
$img = imagecreatefromjpeg($img);
$width = imagesx($img);
$height = imagesy($img);

// Generate an image
//$gdImage = @imagecreatetruecolor($width, $height) or die('Cannot Initialize new GD image stream');
//$textColor = imagecolorallocate($gdImage, 255, 255, 255);
//imagestring($gdImage, 1, 5, 5,  'Created with PhpSpreadsheet', $textColor);

// Add a drawing to the worksheet
$drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
$drawing->setName('SAMPLE ORDER');
$drawing->setDescription('SAMPLE ORDER');
//$drawing->setImageResource($gdImage);
$drawing->setImageResource($img);
$drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
$drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
//$drawing->setHeight($width);

//$drawing->setHeight($width>550 ? 550:$width);
$drawing->setWidth(250);
//$drawing->setHeight(150);
$drawing->setCoordinates('A26');
$drawing->setOffsetX(5);
$drawing->setOffsetY(5);
$drawing->setWorksheet($spreadsheet->getActiveSheet());
/*
$sheet->setCellValue("L18", $pdp1['fab4']); //裁法
$sheet->setCellValue("L22", $pdp1['fab4']); //针距如下
$sheet->setCellValue("L25", $pdp1['fab3']); //工艺说明及注意事项*/





unset($_SESSION['pdp1'] ); //注销SESSION

$output=  ($_GET['action'] == 'formdown' )? 1:0;
//$output= 1;
$filenameout = 'pdp1out.xlsx';
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

    $FILEURL = 'http://office.jmwebseo.cn/highable/output/'.$filenameout;
    $MSFILEURL = 'http://view.officeapps.live.com/op/view.aspx?src='. urlencode($FILEURL);
    //echo "<a href= 'http://view.officeapps.live.com/op/view.aspx?src=". urlencode($FILEURL)."' target='_blank' >跳轉--{$filename}</a>";
    Header("Location:{$MSFILEURL}");
}

