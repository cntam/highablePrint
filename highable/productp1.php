<?php
session_start();

//require '../vendor/autoload.php';
//require '/home/pan/vendor/autoload.php';
require_once('autoloadconfig.php');  //判断是否在线

if($online){
    require_once '/home/pan/vendor/autoload.php';

}else{
    require_once '/Applications/XAMPP/xamppfiles/htdocs/composer/vendor/autoload.php';
}
require_once ('img.php');

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

$sheet->setCellValue('C2',  $productp1['guest']);
$sheet->setCellValue('C3',  $productp1['billdate']);
$sheet->setCellValue('H2',  $productp1['doc']);
$sheet->setCellValue('H3',  $productp1['styleno']);
$sheet->setCellValue('L2',  $productp1['department']);
$sheet->setCellValue('L3',  $productp1['findate']);
$sheet->setCellValue('M3',  $productp1['trans']);


$sheet->setCellValue("C5", $productp1['ct'][5]);
$sheet->setCellValue("D5", $productp1['ct'][6]);
$sheet->setCellValue("E5", $productp1['ct'][7]);
$sheet->setCellValue("F5", $productp1['ct'][8]);
$sheet->setCellValue("G5", $productp1['ct'][9]);
$sheet->setCellValue("H5", $productp1['ct'][10]);
$sheet->setCellValue("I5", $productp1['ct'][11]);
$sheet->setCellValue("J5", $productp1['ct'][12]);


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


$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($productp1['fab1'])) ;
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
$html1 = str_replace('\"', "", htmlspecialchars_decode($productp1['fab2'])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue('L4', '款式图:'.$richText);
$spreadsheet->getActiveSheet()->getStyle('L4')->applyFromArray($styleArray1);

$spreadsheet->getActiveSheet()->getStyle('L4')->getFont()->setSize(8);

$img = $productp1['remarkimg2'];
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
$drawing->setName($productp1['doc']);
$drawing->setDescription($productp1['doc']);
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
$sheet->setCellValue("G".$listrow, $productp1["large"]["o0"]);
$sheet->setCellValue("J".$listrow, $productp1["large"]["o1"]);

$listct = $listrow + 1 ;
$sheet->setCellValue("B".$listct, $productp1['ct'][13]);
$sheet->setCellValue("C".$listct, $productp1['ct'][14]);
$sheet->setCellValue("D".$listct, $productp1['ct'][15]);
$sheet->setCellValue("E".$listct, $productp1['ct'][16]);
$sheet->setCellValue("F".$listct, $productp1['ct'][17]);
$sheet->setCellValue("G".$listct, $productp1['ct'][18]);
$sheet->setCellValue("H".$listct, $productp1['ct'][19]);
$sheet->setCellValue("I".$listct, $productp1['ct'][20]);


$listrowmarker = $listrow + 17  ;
$sheet->setCellValue("A".$listrowmarker, '制单人'); //制单人
$sheet->setCellValue("B".$listrowmarker, $productp1['marker']); //制单人

if($productp1["formnumb"] > 14){    //如果行数大于12 增加行
    $addlist = $listct + 15 ;

    for($n = 1;$n<=($productp1["formnumb"] - 14);$n++ ){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($addlist, 1);
    }
}



$listct = $listct + 1 ;
$formarr = array('A','B','C','D','E','F','G','H','I','J','K');
for($x = 0 ,$c = 1; $c <= count($formarr); $x++ ,$c++){
    $f19 = $listct;
    for($i = 1,$y = 0; $i <= $productp1["formnumb"] ; $i++ ,$y++){
        $sheet->setCellValue($formarr[$x].$f19,  $productp1['large']['c'.$c][$y]);
        $f19++;

    }


}

/*针距如下  */
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($productp1['fab5'])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue("L".($listct-3), $richText);
/*//针距如下 */

/*裁法： */
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($productp1['fab4'])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue("L".($listct-7), $richText);
/*//裁法： */



/*工艺说明及注意事项： */
$wizard = new HtmlHelper();
$html1 = str_replace('\"', "", htmlspecialchars_decode($productp1['fab3'])) ;
$richText = $wizard->toRichTextObject($html1);

$spreadsheet->getActiveSheet() ->setCellValue("L".($listct), '工艺说明及注意事项:  '.$richText);
/*//工艺说明及注意事项： */



unset($_SESSION['productp1'] ); //注销SESSION

//$spreadsheet->getActiveSheet()->getPageMargins()->setRight(0.1); //设置打印边距
//$spreadsheet->getActiveSheet()->getPageMargins()->setLeft(0.1); //*/

$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

$output=  ($_GET['action'] == 'formdown' )? 1:0;
//$output= 1;
$nt = date("YmdHis",time()); //转换为日期。
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

    $FILEURL = 'http://allinone321.com/highable/output/'.$filenameout;
    $MSFILEURL = 'http://view.officeapps.live.com/op/view.aspx?src='. urlencode($FILEURL);

    Header("Location:{$MSFILEURL}");
}
exit();

