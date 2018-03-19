<?php
session_start();

//require '/home/soft/vendor/autoload.php';
//require '../vendor/autoload.php';
require '/home/pan/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器


$samplep4 =   $_SESSION['samplep4'];
//var_dump($samplep4);

//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/samplep4.xlsx');
$sheet = $spreadsheet->getActiveSheet();

$spreadsheet->getActiveSheet()->setTitle("sheet1");

$spreadsheet->getDefaultStyle()->getFont()->setName('微软雅黑');
$spreadsheet->getDefaultStyle()->getFont()->setSize(12);
//$spreadsheet->getActiveSheet()->getDefaultRowDimension()->setRowHeight(50);
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


//$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(15);  //列宽度


//$spreadsheet->getActiveSheet()->getRowDimension('1')->setRowHeight(40); //列高度



// Set cell A1 with a string value

//$spreadsheet->getActiveSheet()->getStyle('A1:D1')->getFont()->setSize(14);


$spreadsheet->getActiveSheet()->setCellValue('B2', $samplep4["category"]);
$spreadsheet->getActiveSheet()->setCellValue('B3', $samplep4["stylename"]);

/*加載圖片*/
$img = $samplep4["logo"];
$img = imagecreatefromjpeg($img);
$width = imagesx($img);
$height = imagesy($img);


// Generate an image
//$gdImage = @imagecreatetruecolor($width, $height) or die('Cannot Initialize new GD image stream');
//$textColor = imagecolorallocate($gdImage, 255, 255, 255);
//imagestring($gdImage, 1, 5, 5,  'Created with PhpSpreadsheet', $textColor);

// Add a drawing to the worksheet
$drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
$drawing->setName($samplep4["stylename"]);
$drawing->setDescription($samplep4["stylename"]);
//$drawing->setImageResource($gdImage);
$drawing->setImageResource($img);
$drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
$drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
//$drawing->setHeight($width);

$drawing->setHeight($height>55 ? 55:$height);
//$drawing->setWidth(180);
//$drawing->setHeight(150);
$drawing->setCoordinates('G2');
$drawing->setOffsetX(5);
$drawing->setOffsetY(5);
$drawing->setWorksheet($spreadsheet->getActiveSheet());

/*加載圖片*/

/*加載圖片*/
$img = $samplep4["remarkimg3"];
$img = imagecreatefromjpeg($img);
$width = imagesx($img);
$height = imagesy($img);


// Generate an image
//$gdImage = @imagecreatetruecolor($width, $height) or die('Cannot Initialize new GD image stream');
//$textColor = imagecolorallocate($gdImage, 255, 255, 255);
//imagestring($gdImage, 1, 5, 5,  'Created with PhpSpreadsheet', $textColor);

// Add a drawing to the worksheet
$drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
$drawing->setName($samplep4["stylename"]);
$drawing->setDescription($samplep4["stylename"]);
//$drawing->setImageResource($gdImage);
$drawing->setImageResource($img);
$drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
$drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
//$drawing->setHeight($width);

$drawing->setWidth($width>285 ? 285:$width);
//$drawing->setWidth(180);
//$drawing->setHeight(150);
$drawing->setCoordinates('A5');
$drawing->setOffsetX(5);
$drawing->setOffsetY(5);
$drawing->setWorksheet($spreadsheet->getActiveSheet());

/*加載圖片*/

/*加載圖片*/
$img = $samplep4["remarkimg4"];
$img = imagecreatefromjpeg($img);
$width = imagesx($img);
$height = imagesy($img);


// Generate an image
//$gdImage = @imagecreatetruecolor($width, $height) or die('Cannot Initialize new GD image stream');
//$textColor = imagecolorallocate($gdImage, 255, 255, 255);
//imagestring($gdImage, 1, 5, 5,  'Created with PhpSpreadsheet', $textColor);

// Add a drawing to the worksheet
$drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
$drawing->setName($samplep4["stylename"]);
$drawing->setDescription($samplep4["stylename"]);
//$drawing->setImageResource($gdImage);
$drawing->setImageResource($img);
$drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
$drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
//$drawing->setHeight($width);

//drawing->setHeight($height>60 ? 60:$height);
$drawing->setWidth($width>285 ? 285:$width);
//$drawing->setHeight(150);
$drawing->setCoordinates('E5');
$drawing->setOffsetX(5);
$drawing->setOffsetY(5);
$drawing->setWorksheet($spreadsheet->getActiveSheet());

/*加載圖片*/


/* 文字模块*/
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($samplep4["remark"]["a1"])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('D5', $richText);
/* 文字模块*/

/* 文字模块*/
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($samplep4["remark"]["b1"])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('D9', $richText);
/* 文字模块*/

/* 文字模块*/
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($samplep4["remark"]["c1"])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('D13', $richText);
/* 文字模块*/

/* 文字模块*/
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($samplep4["remark"]["d1"])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('D17', $richText);
/* 文字模块*/

/* 文字模块*/
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($samplep4["remark"]["e1"])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('D21', $richText);
/* 文字模块*/

/* 文字模块*/
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($samplep4["remark"]["f1"])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('D25', $richText);
/* 文字模块*/

/* 文字模块*/
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($samplep4["remark"]["g1"])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('D29', $richText);
/* 文字模块*/

/* 文字模块*/
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($samplep4["remark"]["h1"])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('D33', $richText);
/* 文字模块*/

/* 文字模块*/
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($samplep4["remark"]["i1"])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('D37', $richText);
/* 文字模块*/

$styleArray1 = [
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
        'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP,
        'wrapText' => true,
        'ShrinkToFit'=>true,
    ],
    'font' => [
        'Size' => '8',
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
$spreadsheet->getActiveSheet()->getStyle("D5:D8")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle("D9:D12")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle("D13:D16")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle("D17:D20")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle("D21:D24")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle("D25:D28")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle("D29:D32")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle("D33:D37")->applyFromArray($styleArray1);

/*右側*/
$spreadsheet->getActiveSheet()->setCellValue('H5', $samplep4["title"]);
$spreadsheet->getActiveSheet()->setCellValue('I6', $samplep4["pattren"]);
$spreadsheet->getActiveSheet()->setCellValue('I7', $samplep4["proto"]);
$spreadsheet->getActiveSheet()->setCellValue('I8', $samplep4["finishingsample"]);
$spreadsheet->getActiveSheet()->setCellValue('I9', $samplep4["referencegarment"]);

/* 文字模块*/
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($samplep4["measurements"])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('H11', $richText);
/* 文字模块*/

/* 文字模块*/
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($samplep4["components"])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('A42', $richText);
/* 文字模块*/

/* 文字模块*/
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($samplep4["notes"])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('E42', $richText);
/* 文字模块*/

$spreadsheet->getActiveSheet()->getStyle("D25:D28")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle("D29:D32")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle("D33:D37")->applyFromArray($styleArray1);



$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$spreadsheet->setActiveSheetIndex(0);

unset($_SESSION['samplep4'] ); //注销SESSION

$output=  ($_GET['action'] == 'formdown' )? 1:0;
$nt = date("YmdHis",time()); //转换为日期。
$filenameout = 'samplep4out'.$nt.'.xlsx';
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
