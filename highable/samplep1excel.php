<?php
session_start();

/*
$client = $samplep1['client'];
$maker = $samplep1['maker'];

$samtime = $samplep1['samtime'];
$pages = $samplep1['pages'];
$clientno = $samplep1['clientno'];
$ordernum = $samplep1['ordernum'];
$transtime1 = $samplep1['transtime1'];
$season = $samplep1['season'];
$cate = $samplep1['cate'];
$filerefer = $samplep1['filerefer'];
$quotas = $samplep1['quotas'];
$transterms = $samplep1['transterms'];
$transmode = $samplep1['transmode'];
$transtime2 = $samplep1['transtime2'];
$refer = $samplep1['refer'];
$styleno = $samplep1['styleno'];
$num = $samplep1['num'];
$client2 = $samplep1['client2'];
$transtime3 = $samplep1['transtime3'];
$sku = $samplep1['sku'];
$samtype = $samplep1['samtype'];
$skucate = $samplep1['skucate'];
$orderremark = $samplep1['orderremark'];
$item = $samplep1['item'];
$material = $samplep1['material'];
$samexplain = $samplep1['samexplain'];
$remark1 = $samplep1['remark1'];
$remarkimg1 = $samplep1['remarkimg1'];
$remark2 = $samplep1['remark2'];
$remarkimg2 = $samplep1['remarkimg2'];
$remark3 = $samplep1['remark3'];
$remarkimg3 = $samplep1['remarkimg3'];
$remark4 = $samplep1['remark4'];
$remarkimg4 = $samplep1['remarkimg4'];

*/
require '/home/pan/vendor/autoload.php';

//require '../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器




$samplep1 =   $_SESSION['samplep1'];
//var_dump($samplep1);

//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/samplep1.xlsx');
$sheet = $spreadsheet->getActiveSheet();

$spreadsheet->getActiveSheet()->setTitle("sheet1");

$spreadsheet->getDefaultStyle()->getFont()->setName('微软雅黑');
$spreadsheet->getDefaultStyle()->getFont()->setSize(12);
//$spreadsheet->getActiveSheet()->getDefaultRowDimension()->setRowHeight(50);

$spreadsheet->getActiveSheet()->getPageMargins()->setRight(0.1);
$spreadsheet->getActiveSheet()->getPageMargins()->setLeft(0.1); //页边距

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
//$spreadsheet->getActiveSheet()->getStyle('A1:D1')->applyFromArray($styleArray1);
/*
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
*/


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

//$spreadsheet->getActiveSheet()->getStyle('A2:B2')->applyFromArray($styleArray);


//$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(15);  //列宽度


//$spreadsheet->getActiveSheet()->getRowDimension('1')->setRowHeight(40); //列高度



// Set cell A1 with a string value

//$spreadsheet->getActiveSheet()->getStyle('A1:D1')->getFont()->setSize(14);


$spreadsheet->getActiveSheet()->setCellValue('B2', $samplep1["client"]);

$spreadsheet->getActiveSheet()->setCellValue('B3', $samplep1["maker"]);
$spreadsheet->getActiveSheet()->setCellValue('B5', $samplep1["orderno"]);
$spreadsheet->getActiveSheet()->setCellValue('E5', $samplep1["samtime"]);
$spreadsheet->getActiveSheet()->setCellValue('H5', $samplep1["pages"]);
// Set cell A2 with a numeric value.

//$spreadsheet->getActiveSheet()->getColumnDimension('A')->setAutoSize(true);
$spreadsheet->getActiveSheet()->setCellValue('B7', $samplep1["clientno"]);

$spreadsheet->getActiveSheet()->setCellValue('E7', $samplep1["ordernum"]);
$spreadsheet->getActiveSheet()->setCellValue('H7', $samplep1["transtime1"]);

$spreadsheet->getActiveSheet()->setCellValue('B8', $samplep1["season"]);
$spreadsheet->getActiveSheet()->setCellValue('E8', $samplep1["cate"]);
$spreadsheet->getActiveSheet()->setCellValue('H8', $samplep1["filerefer"]);



//$spreadsheet->getActiveSheet()->getStyle('A5:C5')->getAlignment()->setWrapText(true);
$spreadsheet->getActiveSheet()->setCellValue('B10', $samplep1["quotas"]);
$spreadsheet->getActiveSheet()->setCellValue('B11', $samplep1["transterms"]);
$spreadsheet->getActiveSheet()->setCellValue('B12', $samplep1["transtime2"]);


$spreadsheet->getActiveSheet()->setCellValue('B14', $samplep1["styleno"]);
$spreadsheet->getActiveSheet()->setCellValue('B15', $samplep1["client2"]);
$spreadsheet->getActiveSheet()->setCellValue('B16', $samplep1["sku"]);
$spreadsheet->getActiveSheet()->setCellValue('B17', $samplep1["skucate"]);
$spreadsheet->getActiveSheet()->setCellValue('B18', $samplep1["item"]);
$spreadsheet->getActiveSheet()->setCellValue('B19', $samplep1["samexplain"]);


$spreadsheet->getActiveSheet()->setCellValue('F11', $samplep1["transmode"]);
$spreadsheet->getActiveSheet()->setCellValue('F12', $samplep1["refer"]);

$spreadsheet->getActiveSheet()->setCellValue('F14', $samplep1["num"]);
$spreadsheet->getActiveSheet()->setCellValue('F15', $samplep1["transtime3"]);
$spreadsheet->getActiveSheet()->setCellValue('F16', $samplep1["samtype"]);
$spreadsheet->getActiveSheet()->setCellValue('F17', $samplep1["orderremark"]);
$spreadsheet->getActiveSheet()->setCellValue('F18', $samplep1["material"]);

/* 图片模块*/
$img = $samplep1["remarkimg1"];
$img = imagecreatefromjpeg($img);
$width = imagesx($img);
$height = imagesy($img);


// Generate an image
//$gdImage = @imagecreatetruecolor($width, $height) or die('Cannot Initialize new GD image stream');
//$textColor = imagecolorallocate($gdImage, 255, 255, 255);
//imagestring($gdImage, 1, 5, 5,  'Created with PhpSpreadsheet', $textColor);

// Add a drawing to the worksheet
$drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
$drawing->setName('remarkimg1');
$drawing->setDescription('remarkimg1');
//$drawing->setImageResource($gdImage);
$drawing->setImageResource($img);
$drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
$drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
//$drawing->setHeight($width);

$drawing->setHeight($width>170 ? 170:$width);
//$drawing->setWidth(250);
//$drawing->setHeight(150);
$drawing->setCoordinates('A21');
$drawing->setOffsetX(5);
$drawing->setOffsetY(5);
$drawing->setWorksheet($spreadsheet->getActiveSheet());
/* 图片模块*/

/* 文字模块*/
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($samplep1['remark2'])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('F21', $richText);
/* 文字模块*/


/* 图片模块*/
$img = $samplep1["remarkimg3"];
$img = imagecreatefromjpeg($img);
$width = imagesx($img);
$height = imagesy($img);


// Generate an image
//$gdImage = @imagecreatetruecolor($width, $height) or die('Cannot Initialize new GD image stream');
//$textColor = imagecolorallocate($gdImage, 255, 255, 255);
//imagestring($gdImage, 1, 5, 5,  'Created with PhpSpreadsheet', $textColor);

// Add a drawing to the worksheet
$drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
$drawing->setName('remarkimg3');
$drawing->setDescription('remarkimg3');
//$drawing->setImageResource($gdImage);
$drawing->setImageResource($img);
$drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
$drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
//$drawing->setHeight($width);

$drawing->setHeight($width>170 ? 170:$width);
//$drawing->setWidth(250);
//$drawing->setHeight(150);
$drawing->setCoordinates('A31');
$drawing->setOffsetX(5);
$drawing->setOffsetY(5);
$drawing->setWorksheet($spreadsheet->getActiveSheet());
/* 图片模块*/

/* 文字模块*/
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($samplep1['remark4'])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('F31', $richText);
/* 文字模块*/


for($j = 0 ; $j < 5 ; $j++) {

    $col = chr(97 + $j);

    for ($i = 2; $i < 10; $i++) {
        $list = chr(66 + $i);
        $x = 41 + $j ;
        //$arr[ $col. $i] = $_POST[$col . $i];
        $spreadsheet->getActiveSheet()->setCellValue($list.$x, $samplep1["color"][$col. $i]);
    }

}
$spreadsheet->getActiveSheet()->setCellValue('B41', $samplep1["color"]["a1"]);
$spreadsheet->getActiveSheet()->setCellValue('A42', $samplep1["color"]["b1"]);
$spreadsheet->getActiveSheet()->setCellValue('A43', $samplep1["color"]["c1"]);
$spreadsheet->getActiveSheet()->setCellValue('A44', $samplep1["color"]["d1"]);
$spreadsheet->getActiveSheet()->setCellValue('A45', $samplep1["color"]["e1"]);
$spreadsheet->getActiveSheet()->setCellValue('K41', '总计');

$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$spreadsheet->setActiveSheetIndex(0);

unset($_SESSION['samplep1'] ); //注销SESSION

$output=  ($_GET['action'] == 'formdown' )? 1:0;
$nt = date("YmdHis",time()); //转换为日期。
$filenameout = 'samplep1out'.$nt.'.xlsx';
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

    Header("Location:{$MSFILEURL}");
}
exit;
