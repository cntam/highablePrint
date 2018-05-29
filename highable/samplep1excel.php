<?php
session_start();

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


$spreadsheet->getActiveSheet()->setCellValue('B2', $samplep1["client"]);

$spreadsheet->getActiveSheet()->setCellValue('B3', $samplep1["maker"]);

$spreadsheet->getActiveSheet()->setCellValue('D2', '工厂');
$spreadsheet->getActiveSheet()->setCellValue('D3', '通知');
$spreadsheet->getActiveSheet()->setCellValue('E2', $samplep1["factory"]);
$spreadsheet->getActiveSheet()->setCellValue('E3', $samplep1["notice"]);


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


$spreadsheet->getActiveSheet()->setCellValue('H11', $samplep1["transmode"]);
$spreadsheet->getActiveSheet()->setCellValue('H12', $samplep1["refer"]);

$spreadsheet->getActiveSheet()->setCellValue('H14', $samplep1["num"]);
$spreadsheet->getActiveSheet()->setCellValue('H15', $samplep1["transtime3"]);
$spreadsheet->getActiveSheet()->setCellValue('H16', $samplep1["samtype"]);
$spreadsheet->getActiveSheet()->setCellValue('H17', $samplep1["orderremark"]);
$spreadsheet->getActiveSheet()->setCellValue('H18', $samplep1["material"]);

/**
 * 图片模块
 */

$img = $samplep1["remarkimg1"];
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

//$drawing->setHeight($width>550 ? 550:$width);
    $drawing->setHeight($height>130 ? 130:$height);
//$drawing->setHeight(150);


    $drawing->setCoordinates("A21");
    $drawing->setOffsetX(5);
    $drawing->setOffsetY(5);
    $drawing->setWorksheet($spreadsheet->getActiveSheet());
}
/* 图片模块 */
///




/* 文字模块*/
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($samplep1['remark2'])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('F21', $richText);
/* 文字模块*/

/**
 * 图片模块
 */

$img = $samplep1["remarkimg3"];
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

//$drawing->setHeight($width>550 ? 550:$width);
    $drawing->setHeight($height>130 ? 130:$height);
//$drawing->setHeight(150);


    $drawing->setCoordinates("A31");
    $drawing->setOffsetX(5);
    $drawing->setOffsetY(5);
    $drawing->setWorksheet($spreadsheet->getActiveSheet());
}
/* 图片模块 */

/* 文字模块*/
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($samplep1['remark4'])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('F31', $richText);
/* 文字模块*/





if ($samplep1['formnum'] > 5) {

    for ($i = 5, $v = 0,$x = 1; $i < $samplep1['formnum']; $i++, $v++,$x++) {


        $spreadsheet->getActiveSheet()->insertNewRowBefore(40, 9);


        $spreadsheet->getActiveSheet()->mergeCells("A40:D47");
        $spreadsheet->getActiveSheet()->mergeCells("F40:K47");
        $spreadsheet->getActiveSheet()->getStyle("A40:D47")->getAlignment()->setWrapText(true);//自动换行
        $spreadsheet->getActiveSheet()->getStyle("F40:K47")->getAlignment()->setWrapText(true);//自动换行

        /**
         * FR 边框线
         */
        $styleArray = [
            'borders' => [
                'outline' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                    'color' => ['argb' => '00000000'],
                ],
            ],
        ];


        $spreadsheet->getActiveSheet()->getStyle("A40:D47")->applyFromArray($styleArray);
        $spreadsheet->getActiveSheet()->getStyle("F40:K47")->applyFromArray($styleArray);
        /* 边框线  */

        /**
         * 图片模块
         */

        $img = $samplep1["remarkimg5"][$v];
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
            $drawing->setName('img');
            $drawing->setDescription('img');
//$drawing->setImageResource($gdImage);
            $drawing->setImageResource($img);
            $drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
            $drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
//$drawing->setHeight($width);

//$drawing->setHeight($width>550 ? 550:$width);
            $drawing->setHeight($height>180 ? 180:$height);
//$drawing->setHeight(150);


            $drawing->setCoordinates("A40");
            $drawing->setOffsetX(5);
            $drawing->setOffsetY(5);
            $drawing->setWorksheet($spreadsheet->getActiveSheet());
        }
        /* 图片模块 */

        /* 文字模块*/
        $wizard = new HtmlHelper();
        //$html1 = str_replace('\"', "", htmlspecialchars_decode($samplep1["remark5"])) ;
        $html1 = str_replace('\"', "", $samplep1["remark5"][$v]) ;
        $richText = $wizard->toRichTextObject($html1);

        $spreadsheet->getActiveSheet() ->setCellValue('F40', $richText);
        /* 文字模块*/

    }

}


if($samplep1['formnum'] > 5){
    $addrow = $samplep1['formnum'] - 5;
}else{
    $addrow = 0;
}
$addrow = $addrow * 9;


for($j = 0 ; $j < 5 ; $j++) {

    $col = chr(97 + $j);

    for ($i = 2; $i < 10; $i++) {
        $list = chr(66 + $i);
        $x = 41 + $j + $addrow;
        //$arr[ $col. $i] = $_POST[$col . $i];
        $spreadsheet->getActiveSheet()->setCellValue($list.$x, $samplep1["color"][$col. $i]);
    }

}

$spreadsheet->getActiveSheet()->setCellValue('B'.(41 + $addrow), $samplep1["color"]["a1"]);
$spreadsheet->getActiveSheet()->setCellValue('A'.(42 + $addrow), $samplep1["color"]["b1"]);
$spreadsheet->getActiveSheet()->setCellValue('A'.(43 + $addrow), $samplep1["color"]["c1"]);
$spreadsheet->getActiveSheet()->setCellValue('A'.(44 + $addrow), $samplep1["color"]["d1"]);
$spreadsheet->getActiveSheet()->setCellValue('A'.(45 + $addrow), $samplep1["color"]["e1"]);
$spreadsheet->getActiveSheet()->setCellValue('K'.(41 + $addrow), '总计');
//
//$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

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
	
	 $FILEURL = 'http://allinone321.com/highable/output/'.$filenameout;
    $MSFILEURL = 'http://view.officeapps.live.com/op/view.aspx?src='. urlencode($FILEURL);

    Header("Location:{$MSFILEURL}");
}
exit;
