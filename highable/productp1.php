<?php
session_start();
require_once('autoloadconfig.php');  //判断是否在线

require_once ('img.php');


use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory; //工廠保存接口
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器

$productp1 =  $_SESSION['productp1'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/productp1.xlsx');
$sheet = $spreadsheet->getActiveSheet();

$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft Yahei');
$spreadsheet->getDefaultStyle()->getFont()->setSize(12);

$styleArray = [

    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
        'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
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

$styleArraytop = [

    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
        'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP,
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

function isselect($value){
    if ( $value == 'on') {
        $output = '■  ';
    } else {
        $output = '□  ';
    }
    return $output;
}


$sheet->setCellValue('C2',  $productp1['guest']);
$sheet->setCellValue('C3',  $productp1['alist']['a1']);
$sheet->setCellValue('I2',  $productp1['jobno']);
$sheet->setCellValue('I3',  $productp1['styleno']);
$sheet->setCellValue('N2',  $productp1['sampleno']);
$sheet->setCellValue('O3',  $productp1['alist']['a2']);




/**
 *  船头办数量
 */
//底部附加行 remark
$sheet->setCellValue('D15',  $productp1['ctlist']['ct23']);
$spreadsheet->getActiveSheet()->getStyle('D15:L15')->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle('D15:L15')->getAlignment()->setWrapText(true);

for($i=0,$ct=1;$i< 14;$i++,$ct++){
    if($ct == 14){
            $sheet->setCellValue('A16',  '出船头办日期：'.$productp1['ctlist']['ct'.$ct]);
            $spreadsheet->getActiveSheet()->getStyle('A16')->applyFromArray($styleArray);
            $spreadsheet->getActiveSheet()->getStyle('A16')->getAlignment()->setWrapText(true);
    }elseif($ct == 12 or $ct == 13){

        if($ct == 12){ //净重：
            $sheet->setCellValue('A15',  '净重：'.$productp1['ctlist']['ct'.$ct]);
            $spreadsheet->getActiveSheet()->getStyle('C15')->applyFromArray($styleArray);
            $spreadsheet->getActiveSheet()->getStyle('C15')->getAlignment()->setWrapText(true);
        }else{  //毛重：
            $sheet->setCellValue('C15',  '毛重：'.$productp1['ctlist']['ct'.$ct]);
            $spreadsheet->getActiveSheet()->getStyle('F15')->applyFromArray($styleArray);
            $spreadsheet->getActiveSheet()->getStyle('F15')->getAlignment()->setWrapText(true);
        }

    }else{
        $row = 14;

            if($ct == 1){
                $col = chr(64 + $ct); //A
            }else{
                $col = chr(64 + $ct); //B
            }

            $sheet->setCellValue($col.$row,  $productp1['ctlist']['ct'.$ct]);
            $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($styleArray);
            $spreadsheet->getActiveSheet()->getStyle($col.$row)->getAlignment()->setWrapText(true);

    }
}
/**
 *  //船头办数量
 */

/**
 *  工艺
 */
//工艺说明及注意事项
//$wizard = new HtmlHelper();
//$html1 = str_replace('\"', "", $productp1['fablist']['fab5']) ;
//$richText = $wizard->toRichTextObject($html1);
//$spreadsheet->getActiveSheet() ->setCellValue('A22', $richText);
//$spreadsheet->getActiveSheet()->getStyle('A22')->getAlignment()->setWrapText(true);

    //$spreadsheet->getActiveSheet()->mergeCells("B25:H48");
    $spreadsheet->getActiveSheet()->getStyle('A22:R28')->getAlignment()->setWrapText(true);  //在单元格中写入换行符“\ n”（ALT +“Enter”）
    $spreadsheet->getActiveSheet()->getStyle("A22:R28")->applyFromArray($styleArraytop);
    $spreadsheet->getActiveSheet()->setCellValue('A22', $productp1['fablist']['fab5']);

    //评语
//$wizard = new HtmlHelper();
//$html1 = str_replace('\"', "", $productp1['fablist']['fab6']) ;
//$richText = $wizard->toRichTextObject($html1);
//$spreadsheet->getActiveSheet() ->setCellValue('A30', $richText);
//$spreadsheet->getActiveSheet()->getStyle('A30')->getAlignment()->setWrapText(true);

$spreadsheet->getActiveSheet()->getStyle('A30:R53')->getAlignment()->setWrapText(true);  //在单元格中写入换行符“\ n”（ALT +“Enter”）
$spreadsheet->getActiveSheet()->getStyle("A30:R53")->applyFromArray($styleArraytop);
$spreadsheet->getActiveSheet()->setCellValue('A30', $productp1['fablist']['fab6']);

//评语附加
//$wizard = new HtmlHelper();
//$html1 = str_replace('\"', "", $productp1['fablist']['fab7']) ;
//$richText = $wizard->toRichTextObject($html1);
//$spreadsheet->getActiveSheet() ->setCellValue('A54', $richText);
//$spreadsheet->getActiveSheet()->getStyle('A54')->getAlignment()->setWrapText(true);

$spreadsheet->getActiveSheet()->getStyle('A54:R59')->getAlignment()->setWrapText(true);  //在单元格中写入换行符“\ n”（ALT +“Enter”）
$spreadsheet->getActiveSheet()->getStyle("A54:R59")->applyFromArray($styleArraytop);
$spreadsheet->getActiveSheet()->setCellValue('A54', $productp1['fablist']['fab7']);


$sheet->setCellValue('B60',  '制单人:  '.$productp1['alist']['a4']);
$spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle($col.$row)->getAlignment()->setWrapText(true);
/**
 *  //工艺
 */


/**
 *  特殊工艺
 */
$teshu = join(",",$productp1['alist']['a5value']);
$sheet->setCellValue('M6',  $teshu);
$spreadsheet->getActiveSheet()->getStyle('M6:P6')->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle('M6:P6')->getAlignment()->setWrapText(true);

/**
 *  //特殊工艺
 */
///**
// *  一.主唛/烟治唛/产地唛车法：
// */
//$sheet->setCellValue('F7',  $productp1['blist']['b1']);
//$sheet->setCellValue('F17',  $productp1['blist']['b2']);
//$wizard = new HtmlHelper();
//$html1 = str_replace('\"', "", $productp1['blist']['b9']) ;
//$richText = $wizard->toRichTextObject($html1);
//$spreadsheet->getActiveSheet() ->setCellValue('C18', $richText);
//$spreadsheet->getActiveSheet()->getStyle('C18')->getAlignment()->setWrapText(true);
//
/*加載圖片*/
$img = $productp1['alist']['a3'];
if ($img == '') {
    $haveimg = false;  //没有图片

} else {

    $path = $img;
    $pathinfo = pathinfo($path);
    //echo "扩展名：$pathinfo[extension]";

    if ($pathinfo["extension"] == 'pdf') {

        $img = pdficon();
        $haveimg = true;
    } else {
        $haveimg = true;
    }
}


if ($haveimg) {
    preg_match('/.(jpg|gif|bmp|jpeg|png)/i', $img, $imgformat);
    $imgformat = $imgformat[1];
    switch ($imgformat) {
        case "jpg":
        case "jpeg":
            $img = imagecreatefromjpeg($img);
            break;
        case "bmp":
            $img = imagecreatefromwbmp($img);
            break;
        case "gif":
            $img = imagecreatefromgif($img);
            break;
        case "png":
            $img = imagecreatefrompng($img);
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

    $drawing->setWidthAndHeight(170,250);  //设置图片最大宽度 高度
    //$drawing->setWidth($width>250 ? 250:$width);
    //$drawing->setHeight($height > 170 ? 170 : $height);
//$drawing->setHeight(150);

    //$drawing->setCoordinates($cola.'2');
    $drawing->setCoordinates('M7');
    $drawing->setOffsetX(5);
    $drawing->setOffsetY(5);
    $drawing->setWorksheet($spreadsheet->getActiveSheet());
}

/*加載圖片*/


$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", $productp1['fablist']['fab2']) ;
$richText = $wizard->toRichTextObject($html1);
$spreadsheet->getActiveSheet() ->setCellValue('Q4', $richText);
$spreadsheet->getActiveSheet()->getStyle('Q4:R14')->getAlignment()->setWrapText(true);

/**
 * 裁法
 */
$M16 = '单方向：'. isselect($productp1['alist']['a6']).'倒插：'.isselect($productp1['alist']['a7']).'女装：'.isselect($productp1['alist']['a8']);
$sheet->setCellValue('M16',  $M16);
$spreadsheet->getActiveSheet()->getStyle('M16:R16')->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle('M16:R16')->getAlignment()->setWrapText(true);

$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", $productp1['fablist']['fab3']) ;
$richText = $wizard->toRichTextObject($html1);
$spreadsheet->getActiveSheet() ->setCellValue('M17', $richText);
$spreadsheet->getActiveSheet()->getStyle('M17')->getAlignment()->setWrapText(true);

$M19 = '过粘朴机 ：'. isselect($productp1['alist']['a9']);
$sheet->setCellValue('M19',  $M19);
$spreadsheet->getActiveSheet()->getStyle('M19:R19')->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle('M19:R19')->getAlignment()->setWrapText(true);

$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", $productp1['fablist']['fab4']) ;
$richText = $wizard->toRichTextObject($html1);
$spreadsheet->getActiveSheet() ->setCellValue('M20', '针距：'.$richText);
$spreadsheet->getActiveSheet()->getStyle('M20:R20')->getAlignment()->setWrapText(true);
/**
 * //裁法
 */

/**
 * 细数分配表
 */
for($v=0,$ct = 15;$v<8;$v++,$ct++){
    $col = chr(68 + $v); //D

    $sheet->setCellValue($col.'5',  $productp1['ctlist']['ct'.$ct]);
    $spreadsheet->getActiveSheet()->getStyle($col.'5')->applyFromArray($styleArray);
    $spreadsheet->getActiveSheet()->getStyle($col.'5')->getAlignment()->setWrapText(true);
}
/**
 *
 */
$row = 12;
$last = $productp1['allot']['formnum'];
for($h=3;$h<=11;$h++){
    $col = chr(65 + $h); //B
    $sheet->setCellValue($col.$row,  $productp1['allot']['b'.$h][$last]);
    $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($styleArray);
    $spreadsheet->getActiveSheet()->getStyle($col.$row)->getAlignment()->setWrapText(true);
}
/**
 * 总数行
 */

$row = 7;
for($i=0;$i<$productp1['allot']['formnum'];$i++){

    //if($i==$productp1['allot']['formnum']){
    if($i>4){
        //$row--;
        $spreadsheet->getActiveSheet()->insertNewRowBefore($row, 1);

        $spreadsheet->getActiveSheet()->mergeCells("A{$row}:B{$row}");
        for($h=1;$h<=11;$h++) {
            if($h == 1){
                $col = chr(64 + $h); //A
            }else{
                $col = chr(65 + $h); //B
            }
            $sheet->setCellValue($col . $row, $productp1['allot']['b' . $h][$i]);
            $spreadsheet->getActiveSheet()->getStyle($col . $row)->applyFromArray($styleArray);
            $spreadsheet->getActiveSheet()->getStyle($col . $row)->getAlignment()->setWrapText(true);
        }
        $row++;
    }else{
        for($h=1;$h<=11;$h++){
            if($h == 1){
                $col = chr(64 + $h); //A
            }else{
                $col = chr(65 + $h); //B
            }

            $sheet->setCellValue($col.$row,  $productp1['allot']['b'.$h][$i]);
            $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($styleArray);
            $spreadsheet->getActiveSheet()->getStyle($col.$row)->getAlignment()->setWrapText(true);
        }
        $row++;
    }
}

/**
 *  //细数分配表
 */
//
///**
// *  //一.主唛/烟治唛/产地唛车法：
// */
//
///**
// *  二.洗水唛位置
// */
//$sheet->setCellValue('P38',  $productp1['blist']['b5']);
//$spreadsheet->getActiveSheet()->getStyle('P38')->getAlignment()->setWrapText(true);
//$sheet->setCellValue('E40',  $productp1['blist']['b6']);
//$spreadsheet->getActiveSheet()->getStyle('E40')->getAlignment()->setWrapText(true);
/////*加載圖片*/
//$img = $productp1['blist']['b4'];
//if ($img == '') {
//    $haveimg = false;  //没有图片
//
//} else {
//
//    $path = $img;
//    $pathinfo = pathinfo($path);
//    //echo "扩展名：$pathinfo[extension]";
//
//    if ($pathinfo["extension"] == 'pdf') {
//
//        $img = pdficon();
//        $haveimg = true;
//    } else {
//        $haveimg = true;
//    }
//}
//
//
//if ($haveimg) {
//    preg_match('/.(jpg|gif|bmp|jpeg|png)/i', $img, $imgformat);
//    $imgformat = $imgformat[1];
//    switch ($imgformat) {
//        case "jpg":
//        case "jpeg":
//            $img = imagecreatefromjpeg($img);
//            break;
//        case "bmp":
//            $img = imagecreatefromwbmp($img);
//            break;
//        case "gif":
//            $img = imagecreatefromgif($img);
//            break;
//        case "png":
//            $img = imagecreatefrompng($img);
//            break;
//    }
//    $width = imagesx($img);
//    $height = imagesy($img);
//
//
//// Add a drawing to the worksheet
//    $drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
//    $drawing->setName('img');
//    $drawing->setDescription('img');
////$drawing->setImageResource($gdImage);
//    $drawing->setImageResource($img);
//    $drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
//    $drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
////$drawing->setHeight($width);
//
//    //$drawing->setWidth($width>250 ? 250:$width);
//    $drawing->setHeight($height > 300 ? 300 : $height);
////$drawing->setHeight(150);
//
//    //$drawing->setCoordinates($cola.'2');
//    $drawing->setCoordinates('D26');
//    $drawing->setOffsetX(5);
//    $drawing->setOffsetY(5);
//    $drawing->setWorksheet($spreadsheet->getActiveSheet());
//}
///*加載圖片*/
///**
// *  //二.洗水唛位置
// */
//
///**
// *   三.挂牌位置
// */
//$sheet->setCellValue('A43',  $productp1['blist']['b8']);
//$spreadsheet->getActiveSheet()->getStyle('A43')->getAlignment()->setWrapText(true);
//
/////*加載圖片*/
//$img = $productp1['blist']['b7'];
//if ($img == '') {
//    $haveimg = false;  //没有图片
//
//} else {
//
//    $path = $img;
//    $pathinfo = pathinfo($path);
//    //echo "扩展名：$pathinfo[extension]";
//
//    if ($pathinfo["extension"] == 'pdf') {
//
//        $img = pdficon();
//        $haveimg = true;
//    } else {
//        $haveimg = true;
//    }
//}
//
//
//if ($haveimg) {
//    preg_match('/.(jpg|gif|bmp|jpeg|png)/i', $img, $imgformat);
//    $imgformat = $imgformat[1];
//    switch ($imgformat) {
//        case "jpg":
//        case "jpeg":
//            $img = imagecreatefromjpeg($img);
//            break;
//        case "bmp":
//            $img = imagecreatefromwbmp($img);
//            break;
//        case "gif":
//            $img = imagecreatefromgif($img);
//            break;
//        case "png":
//            $img = imagecreatefrompng($img);
//            break;
//    }
//    $width = imagesx($img);
//    $height = imagesy($img);
//
//
//// Add a drawing to the worksheet
//    $drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
//    $drawing->setName('img');
//    $drawing->setDescription('img');
////$drawing->setImageResource($gdImage);
//    $drawing->setImageResource($img);
//    $drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
//    $drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
////$drawing->setHeight($width);
//
//    //$drawing->setWidth($width>250 ? 250:$width);
//    $drawing->setHeight($height > 500 ? 500 : $height);
////$drawing->setHeight(150);
//
//    //$drawing->setCoordinates($cola.'2');
//    $drawing->setCoordinates('D44');
//    $drawing->setOffsetX(5);
//    $drawing->setOffsetY(5);
//    $drawing->setWorksheet($spreadsheet->getActiveSheet());
//}
///*加載圖片*/
///**
// *   //三.挂牌位置
// */

$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

//unset($_SESSION['productp1'] ); //注销SESSION

$output=  ($_GET['action'] == 'formdown' )? 1:0;
//$output= 1;
$nt = date("YmdHis",time()); //转换为日期。
$filenameout = 'productp1'.$nt.'.xlsx';
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