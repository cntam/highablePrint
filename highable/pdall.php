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

$pdall =  $_SESSION['pdall'];
//var_dump($pdp1);

//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/pdp1.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
$spreadsheet->getDefaultStyle()->getFont()->setSize(10);

$sheet->setCellValue('A3',  $pdall['pdp1']["SPL_1_code"]);
$sheet->setCellValue('B3',  $pdall['pdp1']["SPL_1_name"]);
$sheet->setCellValue('C3',  $pdall['pdp1']["SPL_1_country"]);
$sheet->setCellValue('D3',  $pdall['pdp1']["SPL_1_contact"]);
$sheet->setCellValue('E3',  $pdall['pdp1']["SPL_1_address"]);
$sheet->setCellValue('F3',  $pdall['pdp1']["SPL_1_email"].'/'.$pdall['pdp1']["SPL_1_tel"].'/'.$pdall['pdp1']["SPL_1_mobile"].'/'.$pdall['pdp1']["SPL_1_qq"]);
$sheet->setCellValue('G3',  $pdall['pdp1']["SPL_1_goods"]);

$sheet->setCellValue('A4',  $pdall['pdp1']["SPL_2_code"]);
$sheet->setCellValue('B4',  $pdall['pdp1']["SPL_2_name"]);
$sheet->setCellValue('C4',  $pdall['pdp1']["SPL_2_country"]);
$sheet->setCellValue('D4',  $pdall['pdp1']["SPL_2_contact"]);
$sheet->setCellValue('E4',  $pdall['pdp1']["SPL_2_address"]);
$sheet->setCellValue('F4',  $pdall['pdp1']["SPL_2_email"].'/'.$pdall['pdp1']["SPL_2_tel"].'/'.$pdall['pdp1']["SPL_2_mobile"].'/'.$pdall['pdp1']["SPL_2_qq"]);
$sheet->setCellValue('G4',  $pdall['pdp1']["SPL_2_goods"]);

$sheet->setCellValue('F8',  $pdall['pdp1']["FR_date"]);
$sheet->setCellValue('F10',  $pdall['pdp1']["FR_ihkno"]);
$sheet->setCellValue('F12',  $pdall['pdp1']["FR_supplier"]);
$sheet->setCellValue('F14',  $pdall['pdp1']["FR_suppliercode"]);
$sheet->setCellValue('F16',  $pdall['pdp1']["FR_comp"]);
$sheet->setCellValue('F18',  $pdall['pdp1']["FR_width"]);
$sheet->setCellValue('F20',  $pdall['pdp1']["FR_weight"]);
$sheet->setCellValue('F22',  $pdall['pdp1']["FR_remark"]);

$sheet->setCellValue('F26',  $pdall['pdp1']["SO_date"]);
$sheet->setCellValue('F28',  $pdall['pdp1']["SO_category"]);
$sheet->setCellValue('F30',  $pdall['pdp1']["SO_styleno"]);
$sheet->setCellValue('F32',  $pdall['pdp1']["SO_client"]);
$sheet->setCellValue('F34',  $pdall['pdp1']["SO_fabric"]);
$sheet->setCellValue('F36',  $pdall['pdp1']["SO_fabricinfo"]);
$sheet->setCellValue('F38',  $pdall['pdp1']["SO_lining"]);
$sheet->setCellValue('F40',  $pdall['pdp1']["SO_lininginfo"]);
$sheet->setCellValue('F42',  $pdall['pdp1']["SO_trim"]);
$sheet->setCellValue('F44',  $pdall['pdp1']["SO_triminfo"]);
$sheet->setCellValue('F46',  $pdall['pdp1']["SO_remark"]);


$img = $pdall['pdp1']["FR_img"];
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



$img = $pdall['pdp1']["SO_img"];
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
$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

/**第二页
 *
 */
$spreadsheet->setActiveSheetIndex(1);  //設置當前活動表
//$spreadsheet->getActiveSheet()->setTitle("sheet2");
//$sheet->setCellValue('A1', 'Hello World !');
$spreadsheet->getDefaultStyle()->getFont()->setName('微软雅黑');
$spreadsheet->getDefaultStyle()->getFont()->setSize(12);
//$spreadsheet->getActiveSheet()->getDefaultRowDimension()->setRowHeight(50);


$styleArray1 = [
    'font' => [
        'bold' => true,
    ],
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
        'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
    ],

];
$spreadsheet->getActiveSheet()->getStyle('B3')->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle('B4')->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle('B5')->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle('B6')->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle('B7')->applyFromArray($styleArray1);



$styleArray = [

    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
        'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
    ],
    //$spreadsheet->getActiveSheet()->getStyle('A2:C2')->getAlignment()->setShrinkToFit(true);//缩小以适合
    'borders' => [

        'bottom' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
        ],

    ],

];

$spreadsheet->getActiveSheet()->getStyle('D3:G3')->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle('D4:G4')->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle('D5:G5')->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle('D6:G6')->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle('D7:G7')->applyFromArray($styleArray);

//$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(15);

$spreadsheet->getActiveSheet()->getRowDimension('3')->setRowHeight(30);
$spreadsheet->getActiveSheet()->getRowDimension('4')->setRowHeight(30);
$spreadsheet->getActiveSheet()->getRowDimension('5')->setRowHeight(30);
$spreadsheet->getActiveSheet()->getRowDimension('6')->setRowHeight(30);
$spreadsheet->getActiveSheet()->getRowDimension('7')->setRowHeight(30);


// Set cell A1 with a string value

//$spreadsheet->getActiveSheet()->getStyle('A1:D1')->getFont()->setSize(14);

//填数据
$spreadsheet->getActiveSheet()->setCellValue('B3', 'IHK NO.');
$spreadsheet->getActiveSheet()->setCellValue('B4', "SUPPLIER &ART.");
$spreadsheet->getActiveSheet()->setCellValue('B5', "COMPOSITION");
$spreadsheet->getActiveSheet()->setCellValue('B6', "FABRIC WIDTH");
$spreadsheet->getActiveSheet()->setCellValue('B7', "REMARK");

// Set cell A2 with a numeric value.

//$spreadsheet->getActiveSheet()->getColumnDimension('A')->setAutoSize(true);
$spreadsheet->getActiveSheet()->mergeCells('D3:G3');
$spreadsheet->getActiveSheet()->setCellValue('D3', $pdall['pdp2']["ihkno"]);
$spreadsheet->getActiveSheet()->mergeCells('D4:G4');
$spreadsheet->getActiveSheet()->setCellValue('D4', $pdall['pdp2']["supplier"]);

// Set cell A2 with a numeric value
$spreadsheet->getActiveSheet()->mergeCells('D5:G5');
$spreadsheet->getActiveSheet()->setCellValue('D5', $pdall['pdp2']["com"]);
$spreadsheet->getActiveSheet()->mergeCells('D6:G6');
$spreadsheet->getActiveSheet()->setCellValue('D6', $pdall['pdp2']["faw"]);

// Set cell A2 with a numeric value
$spreadsheet->getActiveSheet()->mergeCells('D7:G7');
$spreadsheet->getActiveSheet()->setCellValue('D7', $pdall['pdp2']["remark"]);



//$spreadsheet->getActiveSheet()->getStyle('A2:C2')->getAlignment()->setWrapText(true);//自动换行
//$spreadsheet->getActiveSheet()->getStyle('A2:C2')->getAlignment()->setShrinkToFit(true);//缩小以适合
$spreadsheet->getActiveSheet()->getStyle('D3')->getAlignment()->setShrinkToFit(true);//缩小以适合
$spreadsheet->getActiveSheet()->getStyle('D4')->getAlignment()->setShrinkToFit(true);//缩小以适合
$spreadsheet->getActiveSheet()->getStyle('D5')->getAlignment()->setShrinkToFit(true);//缩小以适合
$spreadsheet->getActiveSheet()->getStyle('D6')->getAlignment()->setShrinkToFit(true);//缩小以适合
$spreadsheet->getActiveSheet()->getStyle('D7')->getAlignment()->setShrinkToFit(true);//缩小以适合


//$img = 'http://www.a.cn/wordpress/wp-content/uploads/2018/02/2506390415_532601864.220x220-20.jpg';

$img = imagecreatefromjpeg($pdall['pdp2']["img"]);

$width = imagesx($img);

$height = imagesy($img);


// Generate an image
//$gdImage = @imagecreatetruecolor($width, $height) or die('Cannot Initialize new GD image stream');
//$textColor = imagecolorallocate($gdImage, 255, 255, 255);
//imagestring($gdImage, 1, 5, 5,  'Created with PhpSpreadsheet', $textColor);

// Add a drawing to the worksheet
$drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
$drawing->setName($pdall['pdp2']["ihkno"]);
$drawing->setDescription($pdall['pdp2']["ihkno"]);
//$drawing->setImageResource($gdImage);
$drawing->setImageResource($img);
$drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
$drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
//$drawing->setHeight($width);

//$drawing->setHeight($width>550 ? 550:$width);
$drawing->setWidth($width>460 ? 460:$width);
$drawing->setCoordinates('B9');
$drawing->setOffsetX(10);
$drawing->setOffsetY(20);
$drawing->setWorksheet($spreadsheet->getActiveSheet());



$styleArray2 = [

    'borders' => [
        'top' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
        ],
        'bottom' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
        ],
        'left' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
        ],
        'right' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
        ],
    ],

];

$spreadsheet->getActiveSheet()->getStyle('B9:G34')->applyFromArray($styleArray2);
$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页


/**第三页
 *
 */
$spreadsheet->setActiveSheetIndex(2);  //設置當前活動表
//$spreadsheet->getActiveSheet()->setTitle("sheet3");
//$sheet->setCellValue('A1', 'Hello World !');
$spreadsheet->getDefaultStyle()->getFont()->setName('微软雅黑');
$spreadsheet->getDefaultStyle()->getFont()->setSize(12);
$spreadsheet->getActiveSheet()->getDefaultRowDimension()->setRowHeight(50);
$styleArray1 = [
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
        'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
    ],

    'borders' => [
        'top' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
        ],

    ],

];
$spreadsheet->getActiveSheet()->getStyle('A1:D1')->applyFromArray($styleArray1);

$spreadsheet->getActiveSheet()->getStyle('A1')
    ->getBorders()->getLEFT()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK);
$spreadsheet->getActiveSheet()->getStyle('A1')
    ->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
$spreadsheet->getActiveSheet()->getStyle('B1')
    ->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);

$spreadsheet->getActiveSheet()->getStyle('D1')
    ->getBorders()->getLEFT()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
$spreadsheet->getActiveSheet()->getStyle('D1')
    ->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK);



$styleArray = [

    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
        'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP,
    ],

    'borders' => [
        'top' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
        ],
        'bottom' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
        ],
        'left' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
        ],
        'right' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
        ],
    ],

];

$spreadsheet->getActiveSheet()->getStyle('A2:B2')->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle('C2:D2')->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle('A3:B3')->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle('C3:D3')->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle('A4:B4')->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle('C4:D4')->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle('A5:B5')->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle('C5:D5')->applyFromArray($styleArray);

$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(15);
$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(19);
$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(15);
$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(19);

$spreadsheet->getActiveSheet()->getRowDimension('1')->setRowHeight(40);
$spreadsheet->getActiveSheet()->getRowDimension('2')->setRowHeight(160);
$spreadsheet->getActiveSheet()->getRowDimension('3')->setRowHeight(160);
$spreadsheet->getActiveSheet()->getRowDimension('4')->setRowHeight(160);
$spreadsheet->getActiveSheet()->getRowDimension('5')->setRowHeight(160);


// Set cell A1 with a string value

//$spreadsheet->getActiveSheet()->getStyle('A1:D1')->getFont()->setSize(14);

$spreadsheet->getActiveSheet()->setCellValue('A1', 'CLIENT:');
$spreadsheet->getActiveSheet()->setCellValue('B1', $pdall['pdp3']["client"]);
$spreadsheet->getActiveSheet()->setCellValue('C1', "DATE:");
$spreadsheet->getActiveSheet()->setCellValue('D1', $pdall['pdp3']["date"]);

// Set cell A2 with a numeric value.

//$spreadsheet->getActiveSheet()->getColumnDimension('A')->setAutoSize(true);
$spreadsheet->getActiveSheet()->mergeCells('A2:B2');
$spreadsheet->getActiveSheet()->setCellValue('A2', $pdall['pdp3']['remark1']);
$spreadsheet->getActiveSheet()->mergeCells('C2:D2');
$spreadsheet->getActiveSheet()->setCellValue('C2', $pdall['pdp3']['remark2']);

// Set cell A2 with a numeric value
$spreadsheet->getActiveSheet()->mergeCells('A3:B3');
$spreadsheet->getActiveSheet()->setCellValue('A3', $pdall['pdp3']['remark3']);
$spreadsheet->getActiveSheet()->mergeCells('C3:D3');
$spreadsheet->getActiveSheet()->setCellValue('C3', $pdall['pdp3']['remark4']);

// Set cell A2 with a numeric value
$spreadsheet->getActiveSheet()->mergeCells('A4:B4');
$spreadsheet->getActiveSheet()->setCellValue('A4', $pdall['pdp3']['remark5']);
$spreadsheet->getActiveSheet()->mergeCells('C4:D4');
$spreadsheet->getActiveSheet()->setCellValue('C4', $pdall['pdp3']['remark6']);

// Set cell A2 with a numeric value
$spreadsheet->getActiveSheet()->mergeCells('A5:B5');
$spreadsheet->getActiveSheet()->setCellValue('A5', $pdall['pdp3']['remark7']);
$spreadsheet->getActiveSheet()->mergeCells('C5:D5');
$spreadsheet->getActiveSheet()->setCellValue('C5', $pdall['pdp3']['remark8']);

$spreadsheet->getActiveSheet()->getStyle('A2:C2')->getAlignment()->setWrapText(true);
$spreadsheet->getActiveSheet()->getStyle('A3:C3')->getAlignment()->setWrapText(true);
$spreadsheet->getActiveSheet()->getStyle('A4:C4')->getAlignment()->setWrapText(true);
$spreadsheet->getActiveSheet()->getStyle('A5:C5')->getAlignment()->setWrapText(true);
$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页



unset($_SESSION['pdall'] ); //注销SESSION

// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$spreadsheet->setActiveSheetIndex(0);

$output=  ($_GET['action'] == 'formdown' )? 1:0;
//$output= 1;
$nt = date("YmdHis",time()); //转换为日期。
$filenameout = 'pdallout'.$nt.'.xlsx';
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

