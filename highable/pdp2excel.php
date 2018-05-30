<?php
session_start();
//require '/home/pan/vendor/autoload.php';
require_once('autoloadconfig.php');  //判断是否在线

if($online){
    require_once '/home/pan/vendor/autoload.php';

}else{
    require_once '/Applications/XAMPP/xamppfiles/htdocs/composer/vendor/autoload.php';
}
require_once ('img.php');

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;


$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

$spreadsheet->getActiveSheet()->setTitle("sheet1");
//$sheet->setCellValue('A1', 'Hello World !');
$spreadsheet->getDefaultStyle()->getFont()->setName('微软雅黑');
$spreadsheet->getDefaultStyle()->getFont()->setSize(12);
//$spreadsheet->getActiveSheet()->getDefaultRowDimension()->setRowHeight(50);

$p2page = $_GET['p2page'];

$frlistcon = $_SESSION['frlistcon'];


$ihkno= $frlistcon[$p2page][2];
$supplier = $frlistcon[$p2page][3];
$com =  $frlistcon[$p2page][5];
$faw = $frlistcon[$p2page][6];
$remark = $frlistcon[$p2page][8];
$img = $frlistcon[$p2page][9];


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
$spreadsheet->getActiveSheet()->setCellValue('D3', "$ihkno");
$spreadsheet->getActiveSheet()->mergeCells('D4:G4');
$spreadsheet->getActiveSheet()->setCellValue('D4', "$supplier");

// Set cell A2 with a numeric value
$spreadsheet->getActiveSheet()->mergeCells('D5:G5');
$spreadsheet->getActiveSheet()->setCellValue('D5', "$com");
$spreadsheet->getActiveSheet()->mergeCells('D6:G6');
$spreadsheet->getActiveSheet()->setCellValue('D6', "$faw");

// Set cell A2 with a numeric value
$spreadsheet->getActiveSheet()->mergeCells('D7:G7');
$spreadsheet->getActiveSheet()->setCellValue('D7', "$remark");



//$spreadsheet->getActiveSheet()->getStyle('A2:C2')->getAlignment()->setWrapText(true);//自动换行
//$spreadsheet->getActiveSheet()->getStyle('A2:C2')->getAlignment()->setShrinkToFit(true);//缩小以适合
$spreadsheet->getActiveSheet()->getStyle('D3')->getAlignment()->setShrinkToFit(true);//缩小以适合
$spreadsheet->getActiveSheet()->getStyle('D4')->getAlignment()->setShrinkToFit(true);//缩小以适合
$spreadsheet->getActiveSheet()->getStyle('D5')->getAlignment()->setShrinkToFit(true);//缩小以适合
$spreadsheet->getActiveSheet()->getStyle('D6')->getAlignment()->setShrinkToFit(true);//缩小以适合
$spreadsheet->getActiveSheet()->getStyle('D7')->getAlignment()->setShrinkToFit(true);//缩小以适合


/**
 * 图片模块
 */

$img = $img;
if ($img == '') {
    $haveimg = false;  //没有图片

} else {

    $path = $img;
    $pathinfo = pathinfo($path);
    //echo "扩展名：$pathinfo[extension]";

    if ($pathinfo['extension'] == 'pdf') {

        $img = pdficon();
        $haveimg = true;
    } else {
        $haveimg = true;
    }
}


if ($haveimg){
    preg_match ('/.(jpg|gif|bmp|jpeg|png)/i', $img, $imgformat);
    $imgformat = $imgformat[1];
    switch ($imgformat)
    {
        case "jpg":
        case "jpeg":
            $img = imagecreatefromjpeg($img);
            break;
        case "bmp":
            $img =  imagecreatefromwbmp($img);
            break;
        case "gif":
            $img =  imagecreatefromgif($img);
            break;
        case "png":
            $img =   imagecreatefrompng($img);
            break;
    }
    $width = imagesx($img);
    $height = imagesy($img);


// Add a drawing to the worksheet
    $drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
    $drawing->setName('FABRIC RECODE');
    $drawing->setDescription('FABRIC RECODE');
//$drawing->setImageResource($gdImage);
    $drawing->setImageResource($img);
    $drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
    $drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
//$drawing->setHeight($width);

$drawing->setHeight($width>420 ? 420:$width);
    //$drawing->setHeight($height>130 ? 130:$height);
//$drawing->setHeight(150);


    $drawing->setCoordinates("B9");
    $drawing->setOffsetX(10);
    $drawing->setOffsetY(20);
    $drawing->setWorksheet($spreadsheet->getActiveSheet());
}
/* 图片模块 */


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

$spreadsheet->getActiveSheet()->getStyle('B9:G38')->applyFromArray($styleArray2);



// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$spreadsheet->setActiveSheetIndex(0);

$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

unset($_SESSION['pdp2'] ); //注销SESSION

$output=  ($_GET['action'] == 'formdown' )? 1:0;
$nt = date("YmdHis",time()); //转换为日期。
$filenameout = 'rdp2out'.$nt.'.xlsx';
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
}