<?php
session_start();
header("Content-type: text/html; charset=utf-8");
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



$productp5 =  $_SESSION['productp5'];

//var_dump($productp5);
//echo $productp5[1][2];
//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/productp5.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
$spreadsheet->getDefaultStyle()->getFont()->setSize(10);

$sheet->setCellValue('C1',  $productp5[0]["title"]);
$spreadsheet->getActiveSheet()->getStyle('C1')->getFont()->setSize(16);
$spreadsheet->getActiveSheet()->getStyle('C1')->getFont()->setBold(true);
$sheet->setCellValue('C2',  $productp5[0]["subhead"]);
$sheet->setCellValue('B3',  $productp5[0]["attendee"]);
$sheet->setCellValue('H3',  $productp5[0]["serial"]);

$sheet->setCellValue('B4',  $productp5[0]["styleno"]);
$sheet->setCellValue('D4',  $productp5["doc"]);
$sheet->setCellValue('F4',  $productp5[0]["num"]);
$sheet->setCellValue('H4',  $productp5[0]["atdate"]);
$sheet->setCellValue('B5',  $productp5[0]["style"]);
$sheet->setCellValue('D5',  $productp5[0]["deldate"]);
$sheet->setCellValue('F5',  $productp5[0]["comdate"]);

/*$sheet->setCellValue('L3',  $productp5['findate']);
$sheet->setCellValue('M3',  $productp5['trans']);*/

//$formnuma= $productp5["formnum"] +7;

$listrow = 7;
$formnuma = $productp5[1][0];
for($i= 0,$x = $listrow ; $i <= $formnuma ; $i++){
    //$spreadsheet->getActiveSheet()->mergeCells("A{$x}:H{$x}");
    if($formnuma >2 && $i>2 ){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($x, 1);

    }
    $spreadsheet->getActiveSheet()->mergeCells("A{$x}:H{$x}");
    $sheet->setCellValue("A{$x}", $productp5[8][$i]);
    $x++;
}
$listrow = $listrow + $formnuma + 1 ;
if($productp5[1][1] == '1'){
   $radioa = '■ 有';
   $radiob = '□ 无';
}else{
    $radioa = '□ 有';
    $radiob = '■ 无';
}
$sheet->setCellValue('B'.$listrow, $radioa);
$sheet->setCellValue('C'.$listrow, $radiob);


//echo $listrow ;
$sheet->setCellValue('F'.$listrow, $productp5[1][2]); //处理方法

/* 二、车缝注意事项：*/
$listrow = $listrow +  3 ;
//echo $listrow;
$formnuma = $productp5[2][0];
for($i= 0,$x = $listrow ; $i <= $formnuma ; $i++){
    //$spreadsheet->getActiveSheet()->mergeCells("A{$x}:H{$x}");
    if($formnuma >2 && $i>2 ){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($x, 1);

    }
    $spreadsheet->getActiveSheet()->mergeCells("A{$x}:H{$x}");
    $sheet->setCellValue("A{$x}", $productp5[9][$i]);
    $x++;
}
$listrow = $listrow + $formnuma + 1 ;
if($productp5[2][1] == '1'){
    $radioa = '■ 有';
    $radiob = '□ 无';
}else{
    $radioa = '□ 有';
    $radiob = '■ 无';
}
$sheet->setCellValue('B'.$listrow, $radioa);
$sheet->setCellValue('C'.$listrow, $radiob);


//echo $listrow ;
$sheet->setCellValue('F'.$listrow, $productp5[2][2]); //处理方法
/* //二、车缝注意事项：*/


/* 三、尺寸注意事项：：*/
$listrow = $listrow +  3 ;
//echo $listrow;
$formnuma = $productp5[3][0];
for($i= 0,$x = $listrow ; $i <= $formnuma ; $i++){
    //$spreadsheet->getActiveSheet()->mergeCells("A{$x}:H{$x}");
    if($formnuma >2 && $i>2 ){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($x, 1);

    }
    $spreadsheet->getActiveSheet()->mergeCells("A{$x}:H{$x}");
    $sheet->setCellValue("A{$x}", $productp5[10][$i]);
    $x++;
}
$rowadd = $formnuma > 1 ? 1 :2;
$listrow = $listrow + $formnuma + $rowadd ;

$sheet->setCellValue('F'.$listrow, $productp5[3][1]); //处理方法
/* //三、尺寸注意事项：：*/

/* 四、洗水注意事项：*/
$listrow = $listrow +  3 ;
if($productp5[4][1] == '1'){
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

$formnuma = $productp5[4][0];
for($i= 0,$x = $listrow ; $i <= $formnuma ; $i++){
    //$spreadsheet->getActiveSheet()->mergeCells("A{$x}:H{$x}");
    if($formnuma >2 && $i>2 ){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($x, 1);

    }
    $spreadsheet->getActiveSheet()->mergeCells("A{$x}:H{$x}");
    $sheet->setCellValue("A{$x}", $productp5[11][$i]);
    $x++;
}

/* 四、洗水注意事项：*/

/* 五、整烫注意事项：：*/

$listrow = $listrow +  4 ;
//echo $listrow;
$formnuma = $productp5[5][0];
for($i= 0,$x = $listrow ; $i <= $formnuma ; $i++){
    //$spreadsheet->getActiveSheet()->mergeCells("A{$x}:H{$x}");
    if($formnuma >2 && $i>2 ){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($x, 1);

    }
    $spreadsheet->getActiveSheet()->mergeCells("A{$x}:H{$x}");
    $sheet->setCellValue("A{$x}", $productp5[12][$i]);
    $x++;
}

/* //五、整烫注意事项：*/

/* 六、包装注意事项：*/
$rowadd = $formnuma < 3 ? 3 :0;
$listrow = $listrow + $rowadd ;
$listrow = $listrow +  3 ;
//echo $listrow;
$formnuma = $productp5[6][0];
for($i= 0,$x = $listrow ; $i <= $formnuma ; $i++){
    //$spreadsheet->getActiveSheet()->mergeCells("A{$x}:H{$x}");
    if($formnuma >2 && $i>2 ){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($x, 1);

    }
    $spreadsheet->getActiveSheet()->mergeCells("A{$x}:H{$x}");
    $sheet->setCellValue("A{$x}", $productp5[13][$i]);
    $x++;
}
$listrow = $listrow + $formnuma + 1 ;
if($productp5[6][1] == '1'){
    $radioa = '■ 有';
    $radiob = '□ 无';
}else{
    $radioa = '□ 有';
    $radiob = '■ 无';
}
$sheet->setCellValue('B'.$listrow, $radioa);
$sheet->setCellValue('C'.$listrow, $radiob);


//echo $listrow ;
$sheet->setCellValue('F'.$listrow, $productp5[6][2]); //处理方法

/* //六、包装注意事项：*/

/* 七、其他：*/

$listrow = $listrow +  3 ;
//echo $listrow;
$formnuma = $productp5[7][0];
for($i= 0,$x = $listrow ; $i <= $formnuma ; $i++){
    //$spreadsheet->getActiveSheet()->mergeCells("A{$x}:H{$x}");
    if($formnuma >2 && $i>2 ){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($x, 1);

    }
    $spreadsheet->getActiveSheet()->mergeCells("A{$x}:H{$x}");
    $sheet->setCellValue("A{$x}", $productp5[14][$i]);
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
$sheet->setCellValue('B'.$listrow,  $productp5[0]["rename1"]);
$sheet->setCellValue('F'.$listrow,  $productp5[0]["rename2"]);
/* //七、其他*/

$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

unset($_SESSION['productp5'] ); //注销SESSION


$output=  ($_GET['action'] == 'formdown' )? 1:0;
$nt = date("YmdHis",time()); //转换为日期。
$filenameout = 'productp5out'.$nt.'.xlsx';
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