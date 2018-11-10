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


$sheet->setCellValue('C2',  $productp1['guest']);
$sheet->setCellValue('C3',  $productp1['alist']['a1']);
$sheet->setCellValue('I2',  $productp1['jobno']);
$sheet->setCellValue('I3',  $productp1['styleno']);
$sheet->setCellValue('N2',  $productp1['sampleno']);
$sheet->setCellValue('O3',  $productp1['alist']['a2']);

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
///*加載圖片*/
//$img = $productp1['blist']['b3'];
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
//    $drawing->setHeight($height > 170 ? 170 : $height);
////$drawing->setHeight(150);
//
//    //$drawing->setCoordinates($cola.'2');
//    $drawing->setCoordinates('C8');
//    $drawing->setOffsetX(5);
//    $drawing->setOffsetY(5);
//    $drawing->setWorksheet($spreadsheet->getActiveSheet());
//}
//
///*加載圖片*/
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