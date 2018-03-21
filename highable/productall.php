<?php
session_start();

//require '../vendor/autoload.php';
require '/home/pan/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;

$productall =  $_SESSION['productall'];

//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/productall.xlsx');

$spreadsheet->setActiveSheetIndex(0);  //設置當前活動表
$sheet = $spreadsheet->getActiveSheet();
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
$spreadsheet->getDefaultStyle()->getFont()->setSize(10);

$sheet->setCellValue('C2',  $productall['productp1']['guest']);
$sheet->setCellValue('C3',  $productall['productp1']['billdate']);
$sheet->setCellValue('H2',  $productall['doc']);
$sheet->setCellValue('H3',  $productall['productp1']['styleno']);
$sheet->setCellValue('L2',  $productall['productp1']['department']);
$sheet->setCellValue('L3',  $productall['productp1']['findate']);
$sheet->setCellValue('M3',  $productall['productp1']['trans']);

$formnuma= $productall['productp1']["formnum"] +7;
for($i = 7,$a = 0; $i<$formnuma  ;$i++){
    if($formnuma>12 && $i>11 ){
        $y = $i;
        $spreadsheet->getActiveSheet()->insertNewRowBefore($y, 1);

    }
    $sheet->setCellValue("A{$i}", $productall['productp1']['allot']["b1"][$a]);
    $sheet->setCellValue("B{$i}", $productall['productp1']['allot']["b2"][$a]);
    $sheet->setCellValue("C{$i}", $productall['productp1']['allot']["b3"][$a]);
    $sheet->setCellValue("D{$i}", $productall['productp1']['allot']["b4"][$a]);
    $sheet->setCellValue("E{$i}", $productall['productp1']['allot']["b5"][$a]);
    $sheet->setCellValue("F{$i}", $productall['productp1']['allot']["b6"][$a]);
    $sheet->setCellValue("G{$i}", $productall['productp1']['allot']["b7"][$a]);
    $sheet->setCellValue("H{$i}", $productall['productp1']['allot']["b8"][$a]);
    $sheet->setCellValue("I{$i}", $productall['productp1']['allot']["b9"][$a]);
    $sheet->setCellValue("J{$i}", $productall['productp1']['allot']["b10"][$a]);
    $sheet->setCellValue("K{$i}", $productall['productp1']['allot']["b11"][$a]);

   if($i == $formnuma - 1  ) {
       $z = $formnuma;
       $x = $a +1;
       $sheet->setCellValue("C{$z}", $productall['productp1']['allot']["b3"][$x]);
       $sheet->setCellValue("D{$z}", $productall['productp1']['allot']["b4"][$x]);
       $sheet->setCellValue("E{$z}", $productall['productp1']['allot']["b5"][$x]);
       $sheet->setCellValue("F{$z}", $productall['productp1']['allot']["b6"][$x]);
       $sheet->setCellValue("G{$z}", $productall['productp1']['allot']["b7"][$x]);
       $sheet->setCellValue("H{$z}", $productall['productp1']['allot']["b8"][$x]);
       $sheet->setCellValue("I{$z}", $productall['productp1']['allot']["b9"][$x]);
       $sheet->setCellValue("J{$z}", $productall['productp1']['allot']["b10"][$x]);
       $sheet->setCellValue("K{$z}", $productall['productp1']['allot']["b11"][$x]);
   }

    $a++;

}
$listrow = $formnuma;
$listrow = $listrow + 1;
$sheet->setCellValue("F".$listrow, '产前封样:'.$productall['productp1']['bfsample']);
$listrow = ++$listrow;
$sheet->setCellValue("A".$listrow, $productall['productp1']['ct'][1]);
$sheet->setCellValue("B".$listrow, $productall['productp1']['ct'][2]);
$sheet->setCellValue("C".$listrow, $productall['productp1']['ct'][3]);
$sheet->setCellValue("E".$listrow, $productall['productp1']['ct'][4]);
$listrow = $listrow + 1;
$sheet->setCellValue("C".$listrow , $productall['productp1']['weight']);
$sheet->setCellValue("D".++$listrow , $productall['productp1']['ctdate']);

$spreadsheet->getActiveSheet()->getStyle("A".++$listrow)->getAlignment()->setWrapText(true);


$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($productall['productp1']['fab1'])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue("A".$listrow, '办布如下:'.$richText);

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
$spreadsheet->getActiveSheet()->getStyle("A".$listrow)->applyFromArray($styleArray1);

$spreadsheet->getActiveSheet()->getStyle("A".$listrow)->getFont()->setSize(8);


//$sheet->setCellValue("L4", $productp1['fab2']); //款式图标注

$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($productall['productp1']['fab2'])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('L4', '办布如下:'.$richText);
$spreadsheet->getActiveSheet()->getStyle('L4')->applyFromArray($styleArray1);

$spreadsheet->getActiveSheet()->getStyle('L4')->getFont()->setSize(8);
//$sheet->setCellValue("L7", $productp1['remarkimg2']); //款式图remarkimg2
$img = $productall['productp1']['remarkimg2'];
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


// Generate an image
//$gdImage = @imagecreatetruecolor($width, $height) or die('Cannot Initialize new GD image stream');
//$textColor = imagecolorallocate($gdImage, 255, 255, 255);
//imagestring($gdImage, 1, 5, 5,  'Created with PhpSpreadsheet', $textColor);

// Add a drawing to the worksheet
$drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
$drawing->setName($productall['doc']);
$drawing->setDescription($productall['doc']);
//$drawing->setImageResource($gdImage);
$drawing->setImageResource($img);
$drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
$drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
//$drawing->setHeight($width);

//$drawing->setHeight($width>550 ? 550:$width);
$drawing->setWidth(170);
//$drawing->setHeight(150);
$drawing->setCoordinates('L7');
$drawing->setOffsetX(5);
$drawing->setOffsetY(5);
$drawing->setWorksheet($spreadsheet->getActiveSheet());
/*
$sheet->setCellValue("L18", $productall['productp1']['fab4']); //裁法
$sheet->setCellValue("L22", $productp1['fab4']); //针距如下
$sheet->setCellValue("L25", $productp1['fab3']); //工艺说明及注意事项*/

$listrow =  $listrow + 7;
$sheet->setCellValue("G".$listrow, $productall['productp1']["large"]["o0"]);
$sheet->setCellValue("J".$listrow, $productall['productp1']["large"]["o1"]);

$formnumbrow = $productall['productp1']["formnumb"] > 15 ? ($productall['productp1']["formnumb"] - 15) : 0 ;
$listrowmarker = $listrow + 17 + $formnumbrow ;
$sheet->setCellValue("A".$listrowmarker, '制单人'); //制单人
$sheet->setCellValue("B".$listrowmarker, $productall['productp1']['marker']); //制单人


$listrow= 2 +$listrow;

for($i = $listrow , $a=0; $i<($listrow + $productall['productp1']["formnumb"]) ;$i++){
    if($productall['productp1']["formnumb"]>15 && $i>$listrowmarker-1 ){
        $y = $i;
        $spreadsheet->getActiveSheet()->insertNewRowBefore($y, 1);

    }
    $sheet->setCellValue("A{$i}", $productall['productp1']["large"]['p'.$a]);
    $sheet->setCellValue("B{$i}", $productall['productp1']["large"]['q'.$a]);
    $sheet->setCellValue("C{$i}", $productall['productp1']["large"]['r'.$a]);
    $sheet->setCellValue("D{$i}", $productall['productp1']["large"]['s'.$a]);
    $sheet->setCellValue("E{$i}", $productall['productp1']["large"]['t'.$a]);
    $sheet->setCellValue("F{$i}", $productall['productp1']["large"]['u'.$a]);
    $sheet->setCellValue("G{$i}", $productall['productp1']["large"]['v'.$a]);
    $sheet->setCellValue("H{$i}", $productall['productp1']["large"]['w'.$a]);
    $sheet->setCellValue("I{$i}", $productall['productp1']["large"]['x'.$a]);
    $sheet->setCellValue("J{$i}", $productall['productp1']["large"]['y'.$a]);
    $sheet->setCellValue("K{$i}", $productall['productp1']["large"]['z'.$a]);
    $spreadsheet->getActiveSheet()->getStyle("A{$i}")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->getStyle("B{$i}")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->getStyle("C{$i}")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->getStyle("D{$i}")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->getStyle("E{$i}")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->getStyle("F{$i}")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->getStyle("G{$i}")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->getStyle("H{$i}")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->getStyle("I{$i}")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->getStyle("J{$i}")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->getStyle("K{$i}")->applyFromArray($styleArray1);
    $a++;
}

//$sheet->setCellValue("L".$listrow, '工艺说明及注意事项:  '.$productp1['fab3']); //款式图
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($productall['productp1']['fab3'])) ;
$richText = $wizard->toRichTextObject($html1);

$listrowmarker = $listrowmarker-1; //制单人 行 減1
$spreadsheet->getActiveSheet()->mergeCells("L{$listrow}:N{$listrowmarker}");

$spreadsheet->getActiveSheet() ->setCellValue("L".$listrow, '工艺说明及注意事项:  '.$richText);
//$spreadsheet->getActiveSheet()->getStyle("L".$listrow)->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle("L{$listrow}:N{$listrowmarker}")->applyFromArray($styleArray1);

$spreadsheet->getActiveSheet()->getStyle("L".$listrow)->getFont()->setSize(8);



//$sheet->setCellValue("L".($listrow-3), $productp1['fab5']); //针距如下
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($productall['productp1']['fab5'])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue("L".($listrow-3), $richText);
$cfrow = $listrow - 3;
$spreadsheet->getActiveSheet()->getStyle("L{$cfrow}:N{$listrow}")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle("L{$cfrow}:N{$listrow}")->getFont()->setSize(8);



//$sheet->setCellValue("L".($listrow-7), $productp1['fab4']); //裁法
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($productall['productp1']['fab4'])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue("L".($listrow-7), $richText);

$cfsrow = $listrow-7;
$cferow = $listrow - 4;
$spreadsheet->getActiveSheet()->getStyle("L19:N21")->applyFromArray($styleArray1);


$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页


/**
 * 第二页
 *
 */
$spreadsheet->setActiveSheetIndex(1);  //設置當前活動表
$sheet = $spreadsheet->getActiveSheet();

$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft Yahei');
$spreadsheet->getDefaultStyle()->getFont()->setSize(10);


$sheet->setCellValue('B2',  $productall['productp1']['guest']);
$sheet->setCellValue('B3',  $productall['productp1']['billdate']);
$sheet->setCellValue('D2',  $productall['doc']);
$sheet->setCellValue('D3',  $productall['productp1']['styleno']);
$sheet->setCellValue('F2',  $productall['productp1']['department']);
$sheet->setCellValue('F3',  $productall['productp1']['findate']);
$sheet->setCellValue('G3',  $productall['productp1']['trans']);


$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($productall['productp2']['fab1'])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('A20', $richText);

/*加載圖片*/
$img = $productall['productp2']['remarkimg2'];
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


// Generate an image
//$gdImage = @imagecreatetruecolor($width, $height) or die('Cannot Initialize new GD image stream');
//$textColor = imagecolorallocate($gdImage, 255, 255, 255);
//imagestring($gdImage, 1, 5, 5,  'Created with PhpSpreadsheet', $textColor);

// Add a drawing to the worksheet
$drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
$drawing->setName($productall['doc']);
$drawing->setDescription($productall['doc']);
//$drawing->setImageResource($gdImage);
$drawing->setImageResource($img);
$drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
$drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
//$drawing->setHeight($width);

$drawing->setHeight($height>270 ? 270:$height);
//$drawing->setWidth(180);
//$drawing->setHeight(150);
$drawing->setCoordinates('A5');
$drawing->setOffsetX(5);
$drawing->setOffsetY(5);
$drawing->setWorksheet($spreadsheet->getActiveSheet());

/*加載圖片*/



/*加載圖片*/
$img = $productall['productp2']['remarkimg3'];
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


// Generate an image
//$gdImage = @imagecreatetruecolor($width, $height) or die('Cannot Initialize new GD image stream');
//$textColor = imagecolorallocate($gdImage, 255, 255, 255);
//imagestring($gdImage, 1, 5, 5,  'Created with PhpSpreadsheet', $textColor);

// Add a drawing to the worksheet
$drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
$drawing->setName($productall['doc']);
$drawing->setDescription($productall['doc']);
//$drawing->setImageResource($gdImage);
$drawing->setImageResource($img);
$drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
$drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
//$drawing->setHeight($width);

$drawing->setHeight($height>270 ? 270:$height);
//$drawing->setWidth(180);
//$drawing->setHeight(150);
$drawing->setCoordinates('A27');
$drawing->setOffsetX(5);
$drawing->setOffsetY(5);
$drawing->setWorksheet($spreadsheet->getActiveSheet());

/*加載圖片*/



$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($productall['productp2']['fab2'])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('A42', $richText);


$spreadsheet->getActiveSheet()->getStyle("A20:G25")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle("A42:G48")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

/*第二页*/

/**
 * 第三页
 */
$spreadsheet->setActiveSheetIndex(2);  //設置當前活動表
$sheet = $spreadsheet->getActiveSheet();

$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(15);
$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(16);
$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(52);

$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(16);
$spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(16);
$spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(16);

$sheet->setCellValue('B2',  $productall['doc']);
$sheet->setCellValue('D2',  $productall['productp1']['styleno']);
$sheet->setCellValue('F2',  $productall['productp1']['guest']);



$formnuma= $productall['productp3']["formnum"];
for($i = 0,$a = 0,$row = 4; $i<$formnuma  ;$i++, $row++){
    if($formnuma>25 && $i>24 ){
        $y = $row;
        $spreadsheet->getActiveSheet()->insertNewRowBefore($y, 1);

    }
    $sheet->setCellValue("A{$row}", $productall['productp3']['a1']["a". $a ][0]);
    $spreadsheet->getActiveSheet()->mergeCells("B{$row}:C{$row}");
    $sheet->setCellValue("B{$row}", $productall['productp3']['a1']["b". $a][0]);
    $sheet->setCellValue("D{$row}", $productall['productp3']['a1']["c". $a][0]);
    $sheet->setCellValue("E{$row}", $productall['productp3']['a1']["d". $a][0]);
    $sheet->setCellValue("F{$row}", $productall['productp3']['a1']["e". $a][0]);
    $sheet->setCellValue("G{$row}", $productall['productp3']['a1']["f". $a][0]);

    $spreadsheet->getActiveSheet()->getStyle("A{$row}")->applyFromArray($styleArray1);
    //$spreadsheet->getActiveSheet()->getStyle("B{$row}")->applyFromArray($styleArray1);

    $spreadsheet->getActiveSheet()->getStyle("B{$row}:C{$row}")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->getStyle("D{$row}")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->getStyle("E{$row}")->applyFromArray($styleArray1);

    $spreadsheet->getActiveSheet()->getStyle("F{$row}")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->getStyle("G{$row}")->applyFromArray($styleArray1);
    $a++;

}
$spreadsheet->getActiveSheet()->getPageSetup()
    ->setOrientation(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE); //打印橫向
$spreadsheet->getActiveSheet()->getPageSetup()
    ->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);//打印橫向 A4
$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页
/*第三页*/

/**
 * 第四页
 */
$spreadsheet->setActiveSheetIndex(3);  //設置當前活動表
$sheet = $spreadsheet->getActiveSheet();

$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft Yahei');
$spreadsheet->getDefaultStyle()->getFont()->setSize(12);
$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(15);
for ($v = 1; $v < 19; $v++) {
    $col = chr(97 + $v);
    $spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(9);
}

$sheet->setCellValue('B2',  $productall['productp1']['guest']);
$sheet->setCellValue('B3',  $productall['productp1']['billdate']);
$sheet->setCellValue('H2',  $productall['doc']);
$sheet->setCellValue('H3',  $productall['productp1']['styleno']);
$sheet->setCellValue('P2',  $productall['productp1']['department']);
$sheet->setCellValue('P3',  $productall['productp1']['findate']);
$sheet->setCellValue('R3',  $productall['productp1']['trans']);

$sheet->setCellValue('I4',  $productall['productp4']['a1']["a1"]);
$sheet->setCellValue('M4',  $productall['productp4']['a1']["a2"]);

$formnuma= $productall['productp4']["formnum"] +6;
for($i = 6,$a = 0; $i<$formnuma  ;$i++){
    if($formnuma>12 && $i>11 ){
        $y = $i;
        $spreadsheet->getActiveSheet()->insertNewRowBefore($y, 1);

    }

    $sheet->setCellValue("A{$i}", $productall['productp4']['a1']["b" . $a][0]);
    $sheet->setCellValue("B{$i}", $productall['productp4']['a1']["c". $a][0]);
    $sheet->setCellValue("D{$i}", $productall['productp4']['a1']["d". $a][0]);
    $sheet->setCellValue("F{$i}", $productall['productp4']['a1']["e". $a][0]);
    $sheet->setCellValue("H{$i}", $productall['productp4']['a1']["f". $a][0]);
    $sheet->setCellValue("J{$i}", $productall['productp4']['a1']["g". $a][0]);
    $sheet->setCellValue("L{$i}", $productall['productp4']['a1']["h". $a][0]);
    $sheet->setCellValue("N{$i}", $productall['productp4']['a1']["i". $a][0]);
    $sheet->setCellValue("P{$i}", $productall['productp4']['a1']["j". $a][0]);
    $sheet->setCellValue("R{$i}", $productall['productp4']['a1']["k". $a][0]);
    $sheet->setCellValue("S{$i}", $productall['productp4']['a1']["l". $a][0]);
    for ($v = 0; $v < 19; $v++) {
        $col = chr(97 + $v);
        $spreadsheet->getActiveSheet()->getStyle($col.$i)->applyFromArray($styleArray1);
    }

    $a++;

}
$spreadsheet->getActiveSheet()->getPageSetup()
    ->setOrientation(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE); //打印橫向
$spreadsheet->getActiveSheet()->getPageSetup()
    ->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);//打印橫向 A4
$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

/*第四页*/

/**
 * 第五页
 */
$spreadsheet->setActiveSheetIndex(4);  //設置當前活動表
$sheet = $spreadsheet->getActiveSheet();
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
$spreadsheet->getDefaultStyle()->getFont()->setSize(10);

$sheet->setCellValue('C1',  $productall["productp5"][0]["title"]);
$spreadsheet->getActiveSheet()->getStyle('C1')->getFont()->setSize(16);
$spreadsheet->getActiveSheet()->getStyle('C1')->getFont()->setBold(true);
$sheet->setCellValue('C2',  $productall['productp5'][0]["subhead"]);
$sheet->setCellValue('B3',  $productall['productp5'][0]["attendee"]);
$sheet->setCellValue('H3',  $productall['productp5'][0]["serial"]);

$sheet->setCellValue('B4',  $productall['productp5'][0]["styleno"]);
$sheet->setCellValue('D4',  $productall["doc"]);
$sheet->setCellValue('F4',  $productall['productp5'][0]["num"]);
$sheet->setCellValue('H4',  $productall['productp5'][0]["atdate"]);
$sheet->setCellValue('B5',  $productall['productp5'][0]["style"]);
$sheet->setCellValue('D5',  $productall['productp5'][0]["deldate"]);
$sheet->setCellValue('F5',  $productall['productp5'][0]["comdate"]);

/*$sheet->setCellValue('L3',  $productp5['findate']);
$sheet->setCellValue('M3',  $productp5['trans']);*/

//$formnuma= $productp5["formnum"] +7;

$listrow = 7;
$formnuma = $productall['productp5'][1][0];
for($i= 0,$x = $listrow ; $i <= $formnuma ; $i++){
    //$spreadsheet->getActiveSheet()->mergeCells("A{$x}:H{$x}");
    if($formnuma >2 && $i>2 ){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($x, 1);

    }
    $spreadsheet->getActiveSheet()->mergeCells("A{$x}:H{$x}");
    $sheet->setCellValue("A{$x}", $productall['productp5'][8][$i]);
    $x++;
}
$listrow = $listrow + $formnuma + 1 ;
if($productall['productp5'][1][1] == '1'){
    $radioa = '■ 有';
    $radiob = '□ 无';
}else{
    $radioa = '□ 有';
    $radiob = '■ 无';
}
$sheet->setCellValue('B'.$listrow, $radioa);
$sheet->setCellValue('C'.$listrow, $radiob);


//echo $listrow ;
$sheet->setCellValue('F'.$listrow, $productall['productp5'][1][2]); //处理方法

/* 二、车缝注意事项：*/
$listrow = $listrow +  3 ;
//echo $listrow;
$formnuma = $productall['productp5'][2][0];
for($i= 0,$x = $listrow ; $i <= $formnuma ; $i++){
    //$spreadsheet->getActiveSheet()->mergeCells("A{$x}:H{$x}");
    if($formnuma >2 && $i>2 ){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($x, 1);

    }
    $spreadsheet->getActiveSheet()->mergeCells("A{$x}:H{$x}");
    $sheet->setCellValue("A{$x}", $productall['productp5'][9][$i]);
    $x++;
}
$listrow = $listrow + $formnuma + 1 ;
if($productall['productp5'][2][1] == '1'){
    $radioa = '■ 有';
    $radiob = '□ 无';
}else{
    $radioa = '□ 有';
    $radiob = '■ 无';
}
$sheet->setCellValue('B'.$listrow, $radioa);
$sheet->setCellValue('C'.$listrow, $radiob);


//echo $listrow ;
$sheet->setCellValue('F'.$listrow, $productall['productp5'][2][2]); //处理方法
/* //二、车缝注意事项：*/


/* 三、尺寸注意事项：：*/
$listrow = $listrow +  3 ;
//echo $listrow;
$formnuma = $productall['productp5'][3][0];
for($i= 0,$x = $listrow ; $i <= $formnuma ; $i++){
    //$spreadsheet->getActiveSheet()->mergeCells("A{$x}:H{$x}");
    if($formnuma >2 && $i>2 ){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($x, 1);

    }
    $spreadsheet->getActiveSheet()->mergeCells("A{$x}:H{$x}");
    $sheet->setCellValue("A{$x}", $productall['productp5'][10][$i]);
    $x++;
}
$rowadd = $formnuma > 1 ? 1 :2;
$listrow = $listrow + $formnuma + $rowadd ;

$sheet->setCellValue('F'.$listrow, $productall['productp5'][3][1]); //处理方法
/* //三、尺寸注意事项：：*/

/* 四、洗水注意事项：*/
$listrow = $listrow +  3 ;
if($productall['productp5'][4][1] == '1'){
    $radioa = '■ 需要';
    $radiob = '□ 不需要';
}else{
    $radioa = '□ 需要';
    $radiob = '■ 不需要';
}
$sheet->setCellValue('B'.$listrow, $radioa);
$sheet->setCellValue('C'.$listrow, $radiob);
//echo $listrow;
$listrow = $listrow + 1 ;

$formnuma = $productall['productp5'][4][0];
for($i= 0,$x = $listrow ; $i <= $formnuma ; $i++){
    //$spreadsheet->getActiveSheet()->mergeCells("A{$x}:H{$x}");
    if($formnuma >2 && $i>2 ){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($x, 1);

    }
    $spreadsheet->getActiveSheet()->mergeCells("A{$x}:H{$x}");
    $sheet->setCellValue("A{$x}", $productall['productp5'][11][$i]);
    $x++;
}

/* 四、洗水注意事项：*/

/* 五、整烫注意事项：：*/

$listrow = $listrow +  4 ;
//echo $listrow;
$formnuma = $productall['productp5'][5][0];
for($i= 0,$x = $listrow ; $i <= $formnuma ; $i++){
    //$spreadsheet->getActiveSheet()->mergeCells("A{$x}:H{$x}");
    if($formnuma >2 && $i>2 ){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($x, 1);

    }
    $spreadsheet->getActiveSheet()->mergeCells("A{$x}:H{$x}");
    $sheet->setCellValue("A{$x}", $productall['productp5'][12][$i]);
    $x++;
}

/* //五、整烫注意事项：*/

/* 六、包装注意事项：*/
$rowadd = $formnuma < 3 ? 3 :0;
$listrow = $listrow + $rowadd ;
$listrow = $listrow +  3 ;
//echo $listrow;
$formnuma = $productall['productp5'][6][0];
for($i= 0,$x = $listrow ; $i <= $formnuma ; $i++){
    //$spreadsheet->getActiveSheet()->mergeCells("A{$x}:H{$x}");
    if($formnuma >2 && $i>2 ){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($x, 1);

    }
    $spreadsheet->getActiveSheet()->mergeCells("A{$x}:H{$x}");
    $sheet->setCellValue("A{$x}", $productall['productp5'][13][$i]);
    $x++;
}
$listrow = $listrow + $formnuma + 1 ;
if($productall['productp5'][6][1] == '1'){
    $radioa = '■ 有';
    $radiob = '□ 无';
}else{
    $radioa = '□ 有';
    $radiob = '■ 无';
}
$sheet->setCellValue('B'.$listrow, $radioa);
$sheet->setCellValue('C'.$listrow, $radiob);


//echo $listrow ;
$sheet->setCellValue('F'.$listrow, $productall['productp5'][6][2]); //处理方法

/* //六、包装注意事项：*/

/* 七、其他：*/

$listrow = $listrow +  3 ;
//echo $listrow;
$formnuma = $productall['productp5'][7][0];
for($i= 0,$x = $listrow ; $i <= $formnuma ; $i++){
    //$spreadsheet->getActiveSheet()->mergeCells("A{$x}:H{$x}");
    if($formnuma >2 && $i>2 ){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($x, 1);

    }
    $spreadsheet->getActiveSheet()->mergeCells("A{$x}:H{$x}");
    $sheet->setCellValue("A{$x}", $productall['productp5'][14][$i]);
    $x++;
}
//$rowadd = $formnuma < 3 ? 1 :1;

/*echo $listrow;
echo $formnuma;*/
switch ($formnuma){
    case '0':
        $rowadd = 3;
        break;
    case '1':
        $rowadd = 2;
        break;
    case '2':
        $rowadd = 1;
        break;
    default:
        $rowadd = 1;
}
$listrow = $listrow + $formnuma + $rowadd ;
//echo $listrow;
$sheet->setCellValue('B'.$listrow,  $productall['productp5'][0]["rename1"]);
$sheet->setCellValue('F'.$listrow,  $productall['productp5'][0]["rename2"]);
/* //七、其他*/
$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页
/*第五页*/

/**
 * 第六页
 */
$spreadsheet->setActiveSheetIndex(5);  //設置當前活動表
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

$sheet->setCellValue('B2',  $productall['productp1']['guest']);
$sheet->setCellValue('B3',  $productall['productp1']['billdate']);
$sheet->setCellValue('D2',  $productall['doc']);
$sheet->setCellValue('D3',  $productall['productp1']['styleno']);
$sheet->setCellValue('F2',  $productall['productp1']['department']);
$sheet->setCellValue('F3',  $productall['productp1']['findate']);
$sheet->setCellValue('G3',  $productall['productp1']['trans']);


$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($productall['productp6']['fab1'])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('A5', $richText);




$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($productall['productp6']['fab2'])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('A14', $richText);

$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($productall['productp6']['fab3'])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('A22', $richText);


$spreadsheet->getActiveSheet()->getStyle("A5:G11")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle("A14:G19")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle("A22:G36")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页
/*第六页*/
//unset($_SESSION['productall'] ); //注销SESSION

// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$spreadsheet->setActiveSheetIndex(0);

$output=  ($_GET['action'] == 'formdown' )? 1:0;
//$output= 1;
$nt = date("YmdHis",time()); //转换为日期。
$filenameout = 'productallout'.$nt.'.xlsx';

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
exit();

