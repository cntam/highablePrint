<?php
session_start();
require '../vendor/autoload.php';
//require '/home/pan/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;
$productp3 =  $_SESSION['productp3'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/productp3.xlsx');
$sheet = $spreadsheet->getActiveSheet();

$sheet->setCellValue('B2',  $productp3['doc']);
$sheet->setCellValue('D2',  $productp3['styleno']);
$sheet->setCellValue('F2',  $productp3['guest']);


$formnuma= $productp3["formnum"];
for($i = 0,$a = 0,$row = 4; $i<$formnuma  ;$i++, $row++){
    if($formnuma>25 && $i>24 ){
        $y = $row;
        $spreadsheet->getActiveSheet()->insertNewRowBefore($y, 1);

    }
    $sheet->setCellValue("A{$row}", $productp3['a1']["a" .$a][0]);
    $sheet->setCellValue("B{$row}", $productp3['a1']["b" .$a][0]);
    $sheet->setCellValue("C{$row}", $productp3['a1']["c". $a][0]);
    $sheet->setCellValue("D{$row}", $productp3['a1']["d". $a][0]);
    $sheet->setCellValue("E{$row}", $productp3['a1']["e". $a][0]);
    $sheet->setCellValue("F{$row}", $productp3['a1']["f". $a][0]);


    $a++;

}
/*
$listrow = $formnuma;
$listrow = $listrow + 1;
$sheet->setCellValue("F".$listrow, '产前封样:'.$productp1['bfsample']);
$listrow = ++$listrow;
$sheet->setCellValue("A".$listrow, $productp1['ct'][1]);
$sheet->setCellValue("B".$listrow, $productp1['ct'][2]);
$sheet->setCellValue("C".$listrow, $productp1['ct'][3]);
$sheet->setCellValue("E".$listrow, $productp1['ct'][4]);
$listrow = $listrow + 1;
$sheet->setCellValue("C".$listrow , $productp1['weight']);
$sheet->setCellValue("D".++$listrow , $productp1['ctdate']);

$spreadsheet->getActiveSheet()->getStyle("A".++$listrow)->getAlignment()->setWrapText(true);
$sheet->setCellValue("A".$listrow, '办布如下:'.$productp1['fab1']); //款式图


$sheet->setCellValue("L4", $productp1['fab2']); //款式图标注
//$sheet->setCellValue("L7", $productp1['remarkimg2']); //款式图remarkimg2
$img = $productp1['remarkimg2'];
$img = imagecreatefromjpeg($img);

$width = imagesx($img);

$height = imagesy($img);


// Generate an image
//$gdImage = @imagecreatetruecolor($width, $height) or die('Cannot Initialize new GD image stream');
//$textColor = imagecolorallocate($gdImage, 255, 255, 255);
//imagestring($gdImage, 1, 5, 5,  'Created with PhpSpreadsheet', $textColor);

// Add a drawing to the worksheet
$drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
$drawing->setName($productp1['doc']);
$drawing->setDescription($productp1['doc']);
//$drawing->setImageResource($gdImage);
$drawing->setImageResource($img);
$drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
$drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
//$drawing->setHeight($width);

//$drawing->setHeight($width>550 ? 550:$width);
$drawing->setWidth(180);
//$drawing->setHeight(150);
$drawing->setCoordinates('L7');
$drawing->setOffsetX(5);
$drawing->setOffsetY(5);
$drawing->setWorksheet($spreadsheet->getActiveSheet());
/*
$sheet->setCellValue("L18", $productp1['fab4']); //裁法
$sheet->setCellValue("L22", $productp1['fab4']); //针距如下
$sheet->setCellValue("L25", $productp1['fab3']); //工艺说明及注意事项

$listrow =  $listrow + 7;
$sheet->setCellValue("H".$listrow, $productp1["large"]["o0"]);
$sheet->setCellValue("J".$listrow, $productp1["large"]["o1"]);
$listrow= 2 +$listrow;

for($i = $listrow , $a=0; $i<($listrow + $productp1["formnumb"]) ;$i++){

    $sheet->setCellValue("A{$i}", $productp1["large"]['p'.$a]);
    $sheet->setCellValue("B{$i}", $productp1["large"]['q'.$a]);
    $sheet->setCellValue("C{$i}", $productp1["large"]['r'.$a]);
    $sheet->setCellValue("D{$i}", $productp1["large"]['s'.$a]);
    $sheet->setCellValue("E{$i}", $productp1["large"]['t'.$a]);
    $sheet->setCellValue("F{$i}", $productp1["large"]['u'.$a]);
    $sheet->setCellValue("G{$i}", $productp1["large"]['v'.$a]);
    $sheet->setCellValue("H{$i}", $productp1["large"]['w'.$a]);
    $sheet->setCellValue("I{$i}", $productp1["large"]['x'.$a]);
    $sheet->setCellValue("J{$i}", $productp1["large"]['y'.$a]);
    $sheet->setCellValue("K{$i}", $productp1["large"]['z'.$a]);
    $a++;
}
$sheet->setCellValue("L".$listrow, '工艺说明及注意事项:  '.$productp1['fab3']); //款式图

$sheet->setCellValue("L".($listrow-3), $productp1['fab5']); //针距如下
$sheet->setCellValue("L".($listrow-7), $productp1['fab4']); //针距如下
$spreadsheet->getActiveSheet()->getStyle('L4:N42')->getAlignment()->setWrapText(true);

$sheet->setCellValue("B43", $productp1['marker']); //制单人
*/
//unset($_SESSION['productp3'] ); //注销SESSION

/*$writer = new Xlsx($spreadsheet);
$writer->save('../output/productp3out.xlsx');*/


$output=  ($_GET['action'] == 'formprint' )? 1:0;
$nt = date("YmdHis",time()); //转换为日期。

$filenameout = 'productp3out'.$nt.'.xlsx';
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