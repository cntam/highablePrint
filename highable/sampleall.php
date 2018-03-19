<?php
session_start();

require '/home/pan/vendor/autoload.php';

//require '../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器


$sampleall =   $_SESSION['sampleall'];
//var_dump($samplep1);

//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/sampleall.xlsx');
$sheet = $spreadsheet->getActiveSheet();

$spreadsheet->getActiveSheet()->setTitle("sheet1");

$spreadsheet->getDefaultStyle()->getFont()->setName('微软雅黑');
$spreadsheet->getDefaultStyle()->getFont()->setSize(12);
//$spreadsheet->getActiveSheet()->getDefaultRowDimension()->setRowHeight(50);

$spreadsheet->getActiveSheet()->getPageMargins()->setRight(0.1);
$spreadsheet->getActiveSheet()->getPageMargins()->setLeft(0.1); //页边距

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


$spreadsheet->getActiveSheet()->setCellValue('B2', $sampleall['samplep1']["client"]);

$spreadsheet->getActiveSheet()->setCellValue('B3', $sampleall['samplep1']["maker"]);
$spreadsheet->getActiveSheet()->setCellValue('B5', $sampleall['samplep1']["orderno"]);
$spreadsheet->getActiveSheet()->setCellValue('E5', $sampleall['samplep1']["samtime"]);
$spreadsheet->getActiveSheet()->setCellValue('H5', $sampleall['samplep1']["pages"]);
// Set cell A2 with a numeric value.

//$spreadsheet->getActiveSheet()->getColumnDimension('A')->setAutoSize(true);
$spreadsheet->getActiveSheet()->setCellValue('B7', $sampleall['samplep1']["clientno"]);

$spreadsheet->getActiveSheet()->setCellValue('E7', $sampleall['samplep1']["ordernum"]);
$spreadsheet->getActiveSheet()->setCellValue('H7', $sampleall['samplep1']["transtime1"]);

$spreadsheet->getActiveSheet()->setCellValue('B8', $sampleall['samplep1']["season"]);
$spreadsheet->getActiveSheet()->setCellValue('E8', $sampleall['samplep1']["cate"]);
$spreadsheet->getActiveSheet()->setCellValue('H8', $sampleall['samplep1']["filerefer"]);



//$spreadsheet->getActiveSheet()->getStyle('A5:C5')->getAlignment()->setWrapText(true);
$spreadsheet->getActiveSheet()->setCellValue('B10', $sampleall['samplep1']["quotas"]);
$spreadsheet->getActiveSheet()->setCellValue('B11', $sampleall['samplep1']["transterms"]);
$spreadsheet->getActiveSheet()->setCellValue('B12', $sampleall['samplep1']["transtime2"]);


$spreadsheet->getActiveSheet()->setCellValue('B14', $sampleall['samplep1']["styleno"]);
$spreadsheet->getActiveSheet()->setCellValue('B15', $sampleall['samplep1']["client2"]);
$spreadsheet->getActiveSheet()->setCellValue('B16', $sampleall['samplep1']["sku"]);
$spreadsheet->getActiveSheet()->setCellValue('B17', $sampleall['samplep1']["skucate"]);
$spreadsheet->getActiveSheet()->setCellValue('B18', $sampleall['samplep1']["item"]);
$spreadsheet->getActiveSheet()->setCellValue('B19', $sampleall['samplep1']["samexplain"]);


$spreadsheet->getActiveSheet()->setCellValue('F11', $sampleall['samplep1']["transmode"]);
$spreadsheet->getActiveSheet()->setCellValue('F12', $sampleall['samplep1']["refer"]);

$spreadsheet->getActiveSheet()->setCellValue('F14', $sampleall['samplep1']["num"]);
$spreadsheet->getActiveSheet()->setCellValue('F15', $sampleall['samplep1']["transtime3"]);
$spreadsheet->getActiveSheet()->setCellValue('F16', $sampleall['samplep1']["samtype"]);
$spreadsheet->getActiveSheet()->setCellValue('F17', $sampleall['samplep1']["orderremark"]);
$spreadsheet->getActiveSheet()->setCellValue('F18', $sampleall['samplep1']["material"]);

/* 图片模块*/
$img = $sampleall['samplep1']["remarkimg1"];
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
$html1 = str_replace('\"', "", htmlspecialchars_decode($sampleall['samplep1']['remark2'])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('F21', $richText);
/* 文字模块*/


/* 图片模块*/
$img = $sampleall['samplep1']["remarkimg3"];
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
$html1 = str_replace('\"', "", htmlspecialchars_decode($sampleall['samplep1']['remark4'])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('F31', $richText);
/* 文字模块*/


for($j = 0 ; $j < 5 ; $j++) {

    $col = chr(97 + $j);

    for ($i = 2; $i < 10; $i++) {
        $list = chr(66 + $i);
        $x = 41 + $j ;
        //$arr[ $col. $i] = $_POST[$col . $i];
        $spreadsheet->getActiveSheet()->setCellValue($list.$x, $sampleall['samplep1']["color"][$col. $i]);
    }

}
$spreadsheet->getActiveSheet()->setCellValue('B41', $sampleall['samplep1']["color"]["a1"]);
$spreadsheet->getActiveSheet()->setCellValue('A42', $sampleall['samplep1']["color"]["b1"]);
$spreadsheet->getActiveSheet()->setCellValue('A43', $sampleall['samplep1']["color"]["c1"]);
$spreadsheet->getActiveSheet()->setCellValue('A44', $sampleall['samplep1']["color"]["d1"]);
$spreadsheet->getActiveSheet()->setCellValue('A45', $sampleall['samplep1']["color"]["e1"]);
$spreadsheet->getActiveSheet()->setCellValue('K41', '总计');

$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页


/**
 * 第二页
 */
$spreadsheet->setActiveSheetIndex(1);  //設置當前活動表
$sheet = $spreadsheet->getActiveSheet();

//$spreadsheet->getActiveSheet()->setTitle("sheet2");

$spreadsheet->getDefaultStyle()->getFont()->setName('微软雅黑');
$spreadsheet->getDefaultStyle()->getFont()->setSize(12);


$spreadsheet->getActiveSheet()->setCellValue('B2', $sampleall['samplep1']["client"]);
$spreadsheet->getActiveSheet()->setCellValue('B3', $sampleall['samplep1']["maker"]);
$spreadsheet->getActiveSheet()->setCellValue('B5', $sampleall['samplep1']["orderno"]);
$spreadsheet->getActiveSheet()->setCellValue('E5', $sampleall['samplep1']["samtime"]);
$spreadsheet->getActiveSheet()->setCellValue('H5', $sampleall['samplep1']["pages"]);
// Set cell A2 with a numeric value.


//$spreadsheet->getActiveSheet()->getColumnDimension('A')->setAutoSize(true);
$spreadsheet->getActiveSheet()->setCellValue('B7', $sampleall['samplep1']["clientno"]);

$spreadsheet->getActiveSheet()->setCellValue('E7', $sampleall['samplep1']["ordernum"]);
$spreadsheet->getActiveSheet()->setCellValue('H7', $sampleall['samplep1']["transtime1"]);

$spreadsheet->getActiveSheet()->setCellValue('B8', $sampleall['samplep1']["season"]);
$spreadsheet->getActiveSheet()->setCellValue('E8', $sampleall['samplep1']["cate"]);
$spreadsheet->getActiveSheet()->setCellValue('H8', $sampleall['samplep1']["filerefer"]);



/* 文字模块*/
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($sampleall['samplep2']['fab'])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('C10', $richText);
/* 文字模块*/




/* 文字模块*/
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($sampleall['samplep2']["item"])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('C15', $richText);
/* 文字模块*/

/* 文字模块*/
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($sampleall['samplep2']["comment"])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('C21', $richText);
/* 文字模块*/

/* 文字模块*/
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($sampleall['samplep2']["annotate"])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('C43', $richText);
/* 文字模块*/




$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页
/*第二页*/

/**
 * 第三页
 */
$spreadsheet->setActiveSheetIndex(2);  //設置當前活動表
$sheet = $spreadsheet->getActiveSheet();

//$spreadsheet->getActiveSheet()->setTitle("sheet1");

$spreadsheet->getDefaultStyle()->getFont()->setName('微软雅黑');
$spreadsheet->getDefaultStyle()->getFont()->setSize(12);


$spreadsheet->getActiveSheet()->setCellValue('B2', $sampleall['samplep3']["category"]);
$spreadsheet->getActiveSheet()->setCellValue('B3', $sampleall['samplep3']["stylename"]);

/*加載圖片*/
$img = $sampleall['samplep3']["logo"];
$img = imagecreatefromjpeg($img);
$width = imagesx($img);
$height = imagesy($img);


// Generate an image
//$gdImage = @imagecreatetruecolor($width, $height) or die('Cannot Initialize new GD image stream');
//$textColor = imagecolorallocate($gdImage, 255, 255, 255);
//imagestring($gdImage, 1, 5, 5,  'Created with PhpSpreadsheet', $textColor);

// Add a drawing to the worksheet
$drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
$drawing->setName($sampleall['samplep3']["stylename"]);
$drawing->setDescription($sampleall['samplep3']["stylename"]);
//$drawing->setImageResource($gdImage);
$drawing->setImageResource($img);
$drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
$drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
//$drawing->setHeight($width);

$drawing->setHeight($height>60 ? 60:$height);
//$drawing->setWidth(180);
//$drawing->setHeight(150);
$drawing->setCoordinates('G2');
$drawing->setOffsetX(5);
$drawing->setOffsetY(5);
$drawing->setWorksheet($spreadsheet->getActiveSheet());

/*加載圖片*/

/*加載圖片*/
$img = $sampleall['samplep3']["remarkimg3"];
$img = imagecreatefromjpeg($img);
$width = imagesx($img);
$height = imagesy($img);


// Generate an image
//$gdImage = @imagecreatetruecolor($width, $height) or die('Cannot Initialize new GD image stream');
//$textColor = imagecolorallocate($gdImage, 255, 255, 255);
//imagestring($gdImage, 1, 5, 5,  'Created with PhpSpreadsheet', $textColor);

// Add a drawing to the worksheet
$drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
$drawing->setName($sampleall['samplep3']["stylename"]);
$drawing->setDescription($sampleall['samplep3']["stylename"]);
//$drawing->setImageResource($gdImage);
$drawing->setImageResource($img);
$drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
$drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
//$drawing->setHeight($width);

$drawing->setWidth($width>260 ? 260:$width);
//$drawing->setWidth(180);
//$drawing->setHeight(150);
$drawing->setCoordinates('A6');
$drawing->setOffsetX(5);
$drawing->setOffsetY(5);
$drawing->setWorksheet($spreadsheet->getActiveSheet());

/*加載圖片*/

/*加載圖片*/
$img = $sampleall['samplep3']["remarkimg4"];
$img = imagecreatefromjpeg($img);
$width = imagesx($img);
$height = imagesy($img);


// Generate an image
//$gdImage = @imagecreatetruecolor($width, $height) or die('Cannot Initialize new GD image stream');
//$textColor = imagecolorallocate($gdImage, 255, 255, 255);
//imagestring($gdImage, 1, 5, 5,  'Created with PhpSpreadsheet', $textColor);

// Add a drawing to the worksheet
$drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
$drawing->setName($sampleall['samplep3']["stylename"]);
$drawing->setDescription($sampleall['samplep3']["stylename"]);
//$drawing->setImageResource($gdImage);
$drawing->setImageResource($img);
$drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
$drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
//$drawing->setHeight($width);

//drawing->setHeight($height>60 ? 60:$height);
$drawing->setWidth($width>260 ? 260:$width);
//$drawing->setHeight(150);
$drawing->setCoordinates('G6');
$drawing->setOffsetX(5);
$drawing->setOffsetY(5);
$drawing->setWorksheet($spreadsheet->getActiveSheet());

/*加載圖片*/


/* 文字模块*/
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($sampleall['samplep3']["remark"]["a1"])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('D6', $richText);
/* 文字模块*/




/* 文字模块*/
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($sampleall['samplep3']["remark"]["b1"])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('D13', $richText);
/* 文字模块*/

/* 文字模块*/
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($sampleall['samplep3']["remark"]["c1"])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('D20', $richText);
/* 文字模块*/

/* 文字模块*/
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($sampleall['samplep3']["remark"]["d1"])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('D26', $richText);
/* 文字模块*/

/* 文字模块*/
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($sampleall['samplep3']["remark"]["e1"])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('D32', $richText);
/* 文字模块*/

/* 文字模块*/
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($sampleall['samplep3']["remark"]["f1"])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('D38', $richText);
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

];
$spreadsheet->getActiveSheet()->getStyle("D6:F12")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle("D13:F19")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle("D20:F25")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle("D26:F31")->applyFromArray($styleArray1);

$spreadsheet->getActiveSheet()->getStyle("D26:F32")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle("D32:F38")->applyFromArray($styleArray1);

$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页
/*第三页*/

/**
 * 第四页
 */
$spreadsheet->setActiveSheetIndex(3);  //設置當前活動表
$sheet = $spreadsheet->getActiveSheet();

$spreadsheet->getDefaultStyle()->getFont()->setName('微软雅黑');
$spreadsheet->getDefaultStyle()->getFont()->setSize(12);


$spreadsheet->getActiveSheet()->setCellValue('B2', $sampleall['samplep4']["category"]);
$spreadsheet->getActiveSheet()->setCellValue('B3', $sampleall['samplep4']["stylename"]);

/*加載圖片*/
$img = $sampleall['samplep4']["logo"];
$img = imagecreatefromjpeg($img);
$width = imagesx($img);
$height = imagesy($img);


// Generate an image
//$gdImage = @imagecreatetruecolor($width, $height) or die('Cannot Initialize new GD image stream');
//$textColor = imagecolorallocate($gdImage, 255, 255, 255);
//imagestring($gdImage, 1, 5, 5,  'Created with PhpSpreadsheet', $textColor);

// Add a drawing to the worksheet
$drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
$drawing->setName($sampleall['samplep4']["stylename"]);
$drawing->setDescription($sampleall['samplep4']["stylename"]);
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
$img = $sampleall['samplep4']["remarkimg3"];
$img = imagecreatefromjpeg($img);
$width = imagesx($img);
$height = imagesy($img);


// Generate an image
//$gdImage = @imagecreatetruecolor($width, $height) or die('Cannot Initialize new GD image stream');
//$textColor = imagecolorallocate($gdImage, 255, 255, 255);
//imagestring($gdImage, 1, 5, 5,  'Created with PhpSpreadsheet', $textColor);

// Add a drawing to the worksheet
$drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
$drawing->setName($sampleall['samplep4']["stylename"]);
$drawing->setDescription($sampleall['samplep4']["stylename"]);
//$drawing->setImageResource($gdImage);
$drawing->setImageResource($img);
$drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
$drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
//$drawing->setHeight($width);

$drawing->setWidth($width>250 ? 250:$width);
//$drawing->setWidth(180);
//$drawing->setHeight(150);
$drawing->setCoordinates('A5');
$drawing->setOffsetX(5);
$drawing->setOffsetY(5);
$drawing->setWorksheet($spreadsheet->getActiveSheet());

/*加載圖片*/

/*加載圖片*/
$img = $sampleall['samplep4']["remarkimg4"];
$img = imagecreatefromjpeg($img);
$width = imagesx($img);
$height = imagesy($img);


// Generate an image
//$gdImage = @imagecreatetruecolor($width, $height) or die('Cannot Initialize new GD image stream');
//$textColor = imagecolorallocate($gdImage, 255, 255, 255);
//imagestring($gdImage, 1, 5, 5,  'Created with PhpSpreadsheet', $textColor);

// Add a drawing to the worksheet
$drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
$drawing->setName($sampleall['samplep4']["stylename"]);
$drawing->setDescription($sampleall['samplep4']["stylename"]);
//$drawing->setImageResource($gdImage);
$drawing->setImageResource($img);
$drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
$drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
//$drawing->setHeight($width);

//drawing->setHeight($height>60 ? 60:$height);
$drawing->setWidth($width>250 ? 250:$width);
//$drawing->setHeight(150);
$drawing->setCoordinates('E5');
$drawing->setOffsetX(5);
$drawing->setOffsetY(5);
$drawing->setWorksheet($spreadsheet->getActiveSheet());

/*加載圖片*/


/* 文字模块*/
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($sampleall['samplep4']["remark"]["a1"])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('D5', $richText);
/* 文字模块*/

/* 文字模块*/
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($sampleall['samplep4']["remark"]["b1"])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('D9', $richText);
/* 文字模块*/

/* 文字模块*/
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($sampleall['samplep4']["remark"]["c1"])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('D13', $richText);
/* 文字模块*/

/* 文字模块*/
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($sampleall['samplep4']["remark"]["d1"])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('D17', $richText);
/* 文字模块*/

/* 文字模块*/
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($sampleall['samplep4']["remark"]["e1"])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('D21', $richText);
/* 文字模块*/

/* 文字模块*/
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($sampleall['samplep4']["remark"]["f1"])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('D25', $richText);
/* 文字模块*/

/* 文字模块*/
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($sampleall['samplep4']["remark"]["g1"])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('D29', $richText);
/* 文字模块*/

/* 文字模块*/
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($sampleall['samplep4']["remark"]["h1"])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('D33', $richText);
/* 文字模块*/

/* 文字模块*/
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($sampleall['samplep4']["remark"]["i1"])) ;
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
$spreadsheet->getActiveSheet()->setCellValue('H5', $sampleall['samplep4']["title"]);
$spreadsheet->getActiveSheet()->setCellValue('I6', $sampleall['samplep4']["pattren"]);
$spreadsheet->getActiveSheet()->setCellValue('I7', $sampleall['samplep4']["proto"]);
$spreadsheet->getActiveSheet()->setCellValue('I8', $sampleall['samplep4']["finishingsample"]);
$spreadsheet->getActiveSheet()->setCellValue('I9', $sampleall['samplep4']["referencegarment"]);

/* 文字模块*/
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($sampleall['samplep4']["measurements"])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('H11', $richText);
/* 文字模块*/

/* 文字模块*/
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($sampleall['samplep4']["components"])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('A42', $richText);
/* 文字模块*/

/* 文字模块*/
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($sampleall['samplep4']["notes"])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('E42', $richText);
/* 文字模块*/

$spreadsheet->getActiveSheet()->getStyle("D25:D28")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle("D29:D32")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle("D33:D37")->applyFromArray($styleArray1);



$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页
/*第四页*/

/**
 * 第五页
 */
$spreadsheet->setActiveSheetIndex(4);  //設置當前活動表
$sheet = $spreadsheet->getActiveSheet();

//$spreadsheet->getActiveSheet()->setTitle("sheet1");

$spreadsheet->getDefaultStyle()->getFont()->setName('微软雅黑');
$spreadsheet->getDefaultStyle()->getFont()->setSize(12);

$spreadsheet->getActiveSheet()->setCellValue('B2', $sampleall['samplep5']["category"]);
$spreadsheet->getActiveSheet()->setCellValue('B3', $sampleall['samplep5']["stylename"]);

/*加載圖片*/
$img = $sampleall['samplep5']["logo"];
$img = imagecreatefromjpeg($img);
$width = imagesx($img);
$height = imagesy($img);


// Generate an image
//$gdImage = @imagecreatetruecolor($width, $height) or die('Cannot Initialize new GD image stream');
//$textColor = imagecolorallocate($gdImage, 255, 255, 255);
//imagestring($gdImage, 1, 5, 5,  'Created with PhpSpreadsheet', $textColor);

// Add a drawing to the worksheet
$drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
$drawing->setName($sampleall['samplep5']["stylename"]);
$drawing->setDescription($sampleall['samplep5']["stylename"]);
//$drawing->setImageResource($gdImage);
$drawing->setImageResource($img);
$drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
$drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
//$drawing->setHeight($width);

$drawing->setHeight($height>60 ? 60:$height);
//$drawing->setWidth(180);
//$drawing->setHeight(150);
$drawing->setCoordinates('G2');
$drawing->setOffsetX(5);
$drawing->setOffsetY(5);
$drawing->setWorksheet($spreadsheet->getActiveSheet());

/*加載圖片*/

/*加載圖片*/
$img = $sampleall['samplep5']["remarkimg2"];
$img = imagecreatefromjpeg($img);
$width = imagesx($img);
$height = imagesy($img);


// Generate an image
//$gdImage = @imagecreatetruecolor($width, $height) or die('Cannot Initialize new GD image stream');
//$textColor = imagecolorallocate($gdImage, 255, 255, 255);
//imagestring($gdImage, 1, 5, 5,  'Created with PhpSpreadsheet', $textColor);

// Add a drawing to the worksheet
$drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
$drawing->setName($sampleall['samplep5']["stylename"]);
$drawing->setDescription($sampleall['samplep5']["stylename"]);
//$drawing->setImageResource($gdImage);
$drawing->setImageResource($img);
$drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
$drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
//$drawing->setHeight($width);

$drawing->setWidth($width>420 ? 420:$width);
//$drawing->setWidth(180);
//$drawing->setHeight(150);
$drawing->setCoordinates('A6');
$drawing->setOffsetX(5);
$drawing->setOffsetY(5);
$drawing->setWorksheet($spreadsheet->getActiveSheet());

/*加載圖片*/


/* 文字模块*/
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($sampleall['samplep5']["fab"])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('F5', $richText);
/* 文字模块*/


$styleArray1 = [
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
        'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP,
        'wrapText' => true,
        'ShrinkToFit'=>true,
    ],
    'font' => [
        'Size' => '12',
    ],

];
$spreadsheet->getActiveSheet()->getStyle("F5:I51")->applyFromArray($styleArray1);


$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页
/*第五页*/

// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$spreadsheet->setActiveSheetIndex(0);

unset($_SESSION['$sampleall'] ); //注销SESSION

$output=  ($_GET['action'] == 'formdown' )? 1:0;
$nt = date("YmdHis",time()); //转换为日期。
$filenameout = 'sampleallout'.$nt.'.xlsx';
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
