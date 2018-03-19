<?php
session_start();

//require '../vendor/autoload.php';
require '/home/pan/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;

$productp1 =  $_SESSION['productp1'];

//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/productp1.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
$spreadsheet->getDefaultStyle()->getFont()->setSize(10);

$sheet->setCellValue('C2',  $productp1['guest']);
$sheet->setCellValue('C3',  $productp1['billdate']);
$sheet->setCellValue('H2',  $productp1['doc']);
$sheet->setCellValue('H3',  $productp1['styleno']);
$sheet->setCellValue('L2',  $productp1['department']);
$sheet->setCellValue('L3',  $productp1['findate']);
$sheet->setCellValue('M3',  $productp1['trans']);

$formnuma= $productp1["formnum"] +7;
for($i = 7,$a = 0; $i<$formnuma  ;$i++){
    if($formnuma>12 && $i>11 ){
        $y = $i;
        $spreadsheet->getActiveSheet()->insertNewRowBefore($y, 1);

    }
    $sheet->setCellValue("A{$i}", $productp1['allot']["b1"][$a]);
    $sheet->setCellValue("B{$i}", $productp1['allot']["b2"][$a]);
    $sheet->setCellValue("C{$i}", $productp1['allot']["b3"][$a]);
    $sheet->setCellValue("D{$i}", $productp1['allot']["b4"][$a]);
    $sheet->setCellValue("E{$i}", $productp1['allot']["b5"][$a]);
    $sheet->setCellValue("F{$i}", $productp1['allot']["b6"][$a]);
    $sheet->setCellValue("G{$i}", $productp1['allot']["b7"][$a]);
    $sheet->setCellValue("H{$i}", $productp1['allot']["b8"][$a]);
    $sheet->setCellValue("I{$i}", $productp1['allot']["b9"][$a]);
    $sheet->setCellValue("J{$i}", $productp1['allot']["b10"][$a]);
    $sheet->setCellValue("K{$i}", $productp1['allot']["b11"][$a]);

   if($i == $formnuma - 1  ) {
       $z = $formnuma;
       $x = $a +1;
       $sheet->setCellValue("C{$z}", $productp1['allot']["b3"][$x]);
       $sheet->setCellValue("D{$z}", $productp1['allot']["b4"][$x]);
       $sheet->setCellValue("E{$z}", $productp1['allot']["b5"][$x]);
       $sheet->setCellValue("F{$z}", $productp1['allot']["b6"][$x]);
       $sheet->setCellValue("G{$z}", $productp1['allot']["b7"][$x]);
       $sheet->setCellValue("H{$z}", $productp1['allot']["b8"][$x]);
       $sheet->setCellValue("I{$z}", $productp1['allot']["b9"][$x]);
       $sheet->setCellValue("J{$z}", $productp1['allot']["b10"][$x]);
       $sheet->setCellValue("K{$z}", $productp1['allot']["b11"][$x]);
   }

    $a++;

}
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
//$sheet->setCellValue("A".$listrow, '办布如下:'.str_replace('\"', "", htmlspecialchars_decode($productp1['fab1']))); //款式图

/*
$html1 = str_replace('\"', "", htmlspecialchars_decode($productp1['fab1'])) ;
//$html1 = '办布如下:<b><font size=5>這可怎麼辦？</font></b><div><b><font size=5><strike>很尷尬是吧？</strike></font></b></div><div><b><font size=5><u>就這樣定吧</u></font></b></div>';
*/

$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($productp1['fab1'])) ;
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
        'Size' => '20',
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
$html1 = str_replace('\"', "", htmlspecialchars_decode($productp1['fab2'])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('L4', '办布如下:'.$richText);
$spreadsheet->getActiveSheet()->getStyle('L4')->applyFromArray($styleArray1);

$spreadsheet->getActiveSheet()->getStyle('L4')->getFont()->setSize(8);
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
$sheet->setCellValue("L25", $productp1['fab3']); //工艺说明及注意事项*/

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
//$sheet->setCellValue("L".$listrow, '工艺说明及注意事项:  '.$productp1['fab3']); //款式图
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($productp1['fab3'])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue("L".$listrow, '工艺说明及注意事项:  '.$richText);
$spreadsheet->getActiveSheet()->getStyle("L".$listrow)->applyFromArray($styleArray1);

$spreadsheet->getActiveSheet()->getStyle("L".$listrow)->getFont()->setSize(8);



//$sheet->setCellValue("L".($listrow-3), $productp1['fab5']); //裁法
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($productp1['fab5'])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue("L".($listrow-3), $richText);
$spreadsheet->getActiveSheet()->getStyle("L".($listrow-3))->applyFromArray($styleArray1);

$spreadsheet->getActiveSheet()->getStyle("L".($listrow-3))->getFont()->setSize(8);



//$sheet->setCellValue("L".($listrow-7), $productp1['fab4']); //针距如下
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($productp1['fab4'])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue("L".($listrow-7), $richText);
$spreadsheet->getActiveSheet()->getStyle("L".($listrow-7))->applyFromArray($styleArray1);

$spreadsheet->getActiveSheet()->getStyle("L".($listrow-7))->getFont()->setSize(8);


$spreadsheet->getActiveSheet()->getStyle('L4:N42')->getAlignment()->setWrapText(true);

$sheet->setCellValue("B43", $productp1['marker']); //制单人
unset($_SESSION['productp1'] ); //注销SESSION

$output=  ($_GET['action'] == 'formdown' )? 1:0;
//$output= 1;
$nt=time();
$filenameout = 'productp1out'.$nt.'.xlsx';

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
exit();

