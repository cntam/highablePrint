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

$sheet = $spreadsheet->getActiveSheet();
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
$spreadsheet->getDefaultStyle()->getFont()->setSize(8);
$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(8);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(8);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(8);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(8);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(8);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(8);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('H')->setWidth(8);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('I')->setWidth(8);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('J')->setWidth(8);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('K')->setWidth(8);  //列宽度

$sheet->setCellValue('C2',  $productall['productp1']['guest']);
$sheet->setCellValue('C3',  $productall['productp1']['billdate']);
$sheet->setCellValue('H2',  $productall['productp1']['doc']);
$sheet->setCellValue('H3',  $productall['productp1']['styleno']);
$sheet->setCellValue('L2',  $productall['productp1']['department']);
$sheet->setCellValue('L3',  $productall['productp1']['findate']);
$sheet->setCellValue('M3',  $productall['productp1']['trans']);


$sheet->setCellValue("C5", $productall['productp1']['ct'][5]);
$sheet->setCellValue("D5", $productall['productp1']['ct'][6]);
$sheet->setCellValue("E5", $productall['productp1']['ct'][7]);
$sheet->setCellValue("F5", $productall['productp1']['ct'][8]);
$sheet->setCellValue("G5", $productall['productp1']['ct'][9]);
$sheet->setCellValue("H5", $productall['productp1']['ct'][10]);
$sheet->setCellValue("I5", $productall['productp1']['ct'][11]);
$sheet->setCellValue("J5", $productall['productp1']['ct'][12]);


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




$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($productall['productp1']['fab2'])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('L4', '款式图:'.$richText);
$spreadsheet->getActiveSheet()->getStyle('L4')->applyFromArray($styleArray1);

$spreadsheet->getActiveSheet()->getStyle('L4')->getFont()->setSize(8);

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

$drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
$drawing->setName($productall['productp1']['doc']);
$drawing->setDescription($productall['productp1']['doc']);
//$drawing->setImageResource($gdImage);
$drawing->setImageResource($img);
$drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
$drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
//$drawing->setHeight($width);

//$drawing->setHeight($width>550 ? 550:$width);
$drawing->setWidth(210);
$drawing->setCoordinates('L7');
$drawing->setOffsetX(5);
$drawing->setOffsetY(5);
$drawing->setWorksheet($spreadsheet->getActiveSheet());


$listrow =  $listrow + 7; //行數
$sheet->setCellValue("G".$listrow, $productall['productp1']["large"]["o0"]);
$sheet->setCellValue("J".$listrow, $productall['productp1']["large"]["o1"]);

$listct = $listrow + 1 ;
$sheet->setCellValue("B".$listct, $productall['productp1']['ct'][13]);
$sheet->setCellValue("C".$listct, $productall['productp1']['ct'][14]);
$sheet->setCellValue("D".$listct, $productall['productp1']['ct'][15]);
$sheet->setCellValue("E".$listct, $productall['productp1']['ct'][16]);
$sheet->setCellValue("F".$listct, $productall['productp1']['ct'][17]);
$sheet->setCellValue("G".$listct, $productall['productp1']['ct'][18]);
$sheet->setCellValue("H".$listct, $productall['productp1']['ct'][19]);
$sheet->setCellValue("I".$listct, $productall['productp1']['ct'][20]);


$listrowmarker = $listrow + 17  ;
$sheet->setCellValue("A".$listrowmarker, '制单人'); //制单人
$sheet->setCellValue("B".$listrowmarker, $productall['productp1']['marker']); //制单人

if($productall['productp1']["formnumb"] > 14){    //如果行数大于12 增加行
    $addlist = $listct + 15 ;

    for($n = 1;$n<=($productall['productp1']["formnumb"] - 14);$n++ ){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($addlist, 1);
    }
}



$listct = $listct + 1 ;
$formarr = array('A','B','C','D','E','F','G','H','I','J','K');
for($x = 0 ,$c = 1; $c <= count($formarr); $x++ ,$c++){
    $f19 = $listct;
    for($i = 1,$y = 0; $i <= $productall['productp1']["formnumb"] ; $i++ ,$y++){
        $sheet->setCellValue($formarr[$x].$f19,  $productall['productp1']['large']['c'.$c][$y]);
        $f19++;

    }


}

/*针距如下  */
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($productall['productp1']['fab5'])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue("L".($listct-3), $richText);
/*//针距如下 */

/*裁法： */
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($productall['productp1']['fab4'])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue("L".($listct-7), $richText);
/*//裁法： */



/*工艺说明及注意事项： */
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($productall['productp1']['fab3'])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue("L".($listct), '工艺说明及注意事项:  '.$richText);
/*//工艺说明及注意事项： */


$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页


/**
 * 第二页
 *
 */
$spreadsheet->setActiveSheetIndex(1);  //設置當前活動表
$sheet = $spreadsheet->getActiveSheet();

$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft Yahei');
$spreadsheet->getDefaultStyle()->getFont()->setSize(10);
$styleArray2 = [
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
        'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP,
        'wrapText' => true,
        'ShrinkToFit'=>true,
    ],
    'font' => [
        'Size' => '10',
    ],



];


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


$spreadsheet->getActiveSheet()->getStyle("A20:G25")->applyFromArray($styleArray2);
$spreadsheet->getActiveSheet()->getStyle("A42:G48")->applyFromArray($styleArray2);
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


$sheet->setCellValue('B2', $productall['productp3']['doc']);
$sheet->setCellValue('D2', $productall['productp3']['styleno']);
$sheet->setCellValue('F2', $productall['productp3']['guest']);


$formarr = array('A','B','D','E','F','G');
for($x = 0 ,$c = 1; $c <= count($formarr); $x++ ,$c++){
    $f19 = 4;

    for($i = 1,$y = 0; $i <= $productall['productp3']["formnum"] ; $i++ ,$y++){
        $spreadsheet->getActiveSheet()->mergeCells("B{$f19}:C{$f19}");
        $sheet->setCellValue($formarr[$x].$f19,  $productall['productp3']["a1"]['c'.$c][$y]);
        $spreadsheet->getActiveSheet()->getStyle($formarr[$x].$f19)->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->getStyle("B{$f19}:C{$f19}")->applyFromArray($styleArray1);  //BC样式
        $f19++;

    }
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


$sheet->setCellValue('B2',  $productall['productp4']['guest']);
$sheet->setCellValue('B3',  $productall['productp4']['billdate']);
$sheet->setCellValue('H2',  $productall['doc']);
$sheet->setCellValue('H3',  $productall['productp4']['styleno']);
$sheet->setCellValue('P2',  $productall['productp4']['department']);
$sheet->setCellValue('P3',  $productall['productp4']['findate']);
$sheet->setCellValue('R3',  $productall['productp4']['trans']);

$sheet->setCellValue('I4',  $productall['productp4']['large']['o0']);
$sheet->setCellValue('M4',  $productall['productp4']['large']['o1']);

$sheet->setCellValue("B5", $productall['productp4']['ct'][13]);
$sheet->setCellValue("D5", $productall['productp4']['ct'][14]);
$sheet->setCellValue("F5", $productall['productp4']['ct'][15]);
$sheet->setCellValue("H5", $productall['productp4']['ct'][16]);
$sheet->setCellValue("J5", $productall['productp4']['ct'][17]);
$sheet->setCellValue("L5", $productall['productp4']['ct'][18]);
$sheet->setCellValue("N5", $productall['productp4']['ct'][19]);
$sheet->setCellValue("P5", $productall['productp4']['ct'][20]);



$formarr = array('A','B','D','F','H','J','L','N','P','R','S');
for($x = 0 ,$c = 1; $c <= count($formarr); $x++ ,$c++){
    $f19 = 6;
    for($i = 1,$y = 0; $i <= $productall['productp4']["formnumb"] ; $i++ ,$y++){
        $sheet->setCellValue($formarr[$x].$f19,  $productall['productp4']['large']['c'.$c][$y]);
        $f19++;

    }
}


for($x = 0 ,$c = 1; $c <= 19; $x++ ,$c++){
    $f19 = 6;
    for($i = 1,$y = 0; $i <= $productall['productp4']["formnumb"] ; $i++ ,$y++){
        $col = chr(97 + $x);
        $spreadsheet->getActiveSheet()->getStyle($col.$f19)->applyFromArray($styleArray1);
        $f19++;
    }
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

for ($v = 1; $v <= 8; $v++) {
    $col = chr(97 + $v);
    $spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(11);
}


$sheet->setCellValue('C1',  $productall['productp5'][0]["title"]);
$spreadsheet->getActiveSheet()->getStyle('C1')->getFont()->setSize(16);
$spreadsheet->getActiveSheet()->getStyle('C1')->getFont()->setBold(true);
$sheet->setCellValue('C2',  $productall['productp5'][0]["subhead"]);
$sheet->setCellValue('B3',  $productall['productp5'][0]["attendee"]);
$sheet->setCellValue('H3',  $productall['productp5'][0]["serial"]);

$sheet->setCellValue('B4',  $productall['productp5'][0]["styleno"]);
$sheet->setCellValue('D4',  $productall['doc']);
$sheet->setCellValue('F4',  $productall['productp5'][0]["num"]);
$sheet->setCellValue('H4',  $productall['productp5'][0]["atdate"]);
$sheet->setCellValue('B5',  $productall['productp5'][0]["style"]);
$sheet->setCellValue('D5',  $productall['productp5'][0]["deldate"]);
$sheet->setCellValue('F5',  $productall['productp5'][0]["comdate"]);



$listrow = 7;

$thisrow = $listrow;
for($x = 0 ,$c = 1; $c <= $productall['productp5'][1][0]; $x++ ,$c++){

    if($c >3){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($thisrow, 1);
    }
    $spreadsheet->getActiveSheet()->mergeCells("A{$thisrow}:H{$thisrow}");
    $sheet->setCellValue("A{$thisrow}", $productall['productp5'][8][$x]);
    $thisrow++;

}
$listrow = ($productall['productp5'][1][0]>3) ? ($listrow + $productall['productp5'][1][0]) : ($listrow+3);
//$sheet->setCellValue("L1", $listrow);

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

/** 二、车缝注意事项：*/
$listrow = $listrow +  3 ;

$thisrow = $listrow;
for($x = 0 ,$c = 1; $c <= $productall['productp5'][2][0]; $x++ ,$c++){

    if($c >3){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($thisrow, 1);
    }
    $spreadsheet->getActiveSheet()->mergeCells("A{$thisrow}:H{$thisrow}");
    $sheet->setCellValue("A{$thisrow}", $productall['productp5'][9][$x]);
    $thisrow++;

}
$listrow = ($productall['productp5'][2][0]>3) ? ($listrow + $productall['productp5'][2][0]) : ($listrow+3);
//$sheet->setCellValue("L2", $listrow);

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
///* //二、车缝注意事项：*/



/** 三、尺寸注意事项：：*/
$listrow = $listrow +  3 ;

$thisrow = $listrow;
for($x = 0 ,$c = 1; $c <= $productall['productp5'][3][0]; $x++ ,$c++){

    if($c >3){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($thisrow, 1);
    }
    $spreadsheet->getActiveSheet()->mergeCells("A{$thisrow}:H{$thisrow}");
    $sheet->setCellValue("A{$thisrow}", $productall['productp5'][10][$x]);
    $thisrow++;

}
$listrow = ($productall['productp5'][3][0]>3) ? ($listrow + $productall['productp5'][3][0]) : ($listrow+3);
//$sheet->setCellValue("L3", $listrow);

$sheet->setCellValue('F'.$listrow, $productall['productp5'][3][1]); //处理方法
///* //三、尺寸注意事项：：*/
//
/** 四、洗水注意事项：*/
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

$listrow = $listrow + 1 ;

$thisrow = $listrow;
for($x = 0 ,$c = 1; $c <= $productall['productp5'][4][0]; $x++ ,$c++){

    if($c >3){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($thisrow, 1);
    }
    $spreadsheet->getActiveSheet()->mergeCells("A{$thisrow}:H{$thisrow}");
    $sheet->setCellValue("A{$thisrow}", $productall['productp5'][11][$x]);
    $thisrow++;

}
$listrow = ($productall['productp5'][4][0]>3) ? ($listrow + $productall['productp5'][4][0]) : ($listrow+3);

///* 四、洗水注意事项：*/
//
/** 五、整烫注意事项：：*/
$listrow = $listrow +  2 ;

$thisrow = $listrow;
for($x = 0 ,$c = 1; $c <= $productall['productp5'][5][0]; $x++ ,$c++){

    if($c >3){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($thisrow, 1);
    }
    $spreadsheet->getActiveSheet()->mergeCells("A{$thisrow}:H{$thisrow}");
    $sheet->setCellValue("A{$thisrow}", $productall['productp5'][12][$x]);
    $thisrow++;

}
$listrow = ($productall['productp5'][5][0]>3) ? ($listrow + $productall['productp5'][5][0]) : ($listrow+3);

///* //五、整烫注意事项：*/
//
/** 六、包装注意事项：*/

$listrow = $listrow +  2 ;

$thisrow = $listrow;
for($x = 0 ,$c = 1; $c <= $productall['productp5'][6][0]; $x++ ,$c++){

    if($c >3){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($thisrow, 1);
    }
    $spreadsheet->getActiveSheet()->mergeCells("A{$thisrow}:H{$thisrow}");
    $sheet->setCellValue("A{$thisrow}", $productall['productp5'][13][$x]);
    $thisrow++;

}
$listrow = ($productall['productp5'][6][0]>3) ? ($listrow + $productall['productp5'][6][0]) : ($listrow+3);
//$sheet->setCellValue("L1", $listrow);

if($productall['productp5'][6][1] == '1'){
    $radioa = '■ 有';
    $radiob = '□ 无';
}else{
    $radioa = '□ 有';
    $radiob = '■ 无';
}
$sheet->setCellValue('B'.$listrow, $radioa);
$sheet->setCellValue('C'.$listrow, $radiob);

$sheet->setCellValue('F'.$listrow, $productall['productp5'][6][2]); //处理方法

/* 六、包装注意事项：*/

/* 七、其他：*/
$listrow = $listrow +  3 ;

$thisrow = $listrow;
for($x = 0 ,$c = 1; $c <= $productall['productp5'][7][0]; $x++ ,$c++){

    if($c >3){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($thisrow, 1);
    }
    $spreadsheet->getActiveSheet()->mergeCells("A{$thisrow}:H{$thisrow}");
    $sheet->setCellValue("A{$thisrow}", $productall['productp5'][14][$x]);
    $thisrow++;

}
$listrow = ($productall['productp5'][7][0]>3) ? ($listrow + $productall['productp5'][7][0]) : ($listrow+3);
//$sheet->setCellValue("L1", $listrow);

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
///*第六页*/
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

