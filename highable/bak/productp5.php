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

//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/productp5.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
$spreadsheet->getDefaultStyle()->getFont()->setSize(10);
for ($v = 1; $v <= 8; $v++) {
    $col = chr(97 + $v);
    $spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(11);
}


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


$listrow = 7;

$thisrow = $listrow;
for($x = 0 ,$c = 1; $c <= $productp5[1][0]; $x++ ,$c++){

    if($c >3){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($thisrow, 1);
    }
    $spreadsheet->getActiveSheet()->mergeCells("A{$thisrow}:H{$thisrow}");
    $sheet->setCellValue("A{$thisrow}", $productp5[8][$x]);
    $thisrow++;

}
$listrow = ($productp5[1][0]>3) ? ($listrow + $productp5[1][0]) : ($listrow+3);
//$sheet->setCellValue("L1", $listrow);

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

/** 二、车缝注意事项：*/
$listrow = $listrow +  3 ;

$thisrow = $listrow;
for($x = 0 ,$c = 1; $c <= $productp5[2][0]; $x++ ,$c++){

    if($c >3){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($thisrow, 1);
    }
    $spreadsheet->getActiveSheet()->mergeCells("A{$thisrow}:H{$thisrow}");
    $sheet->setCellValue("A{$thisrow}", $productp5[9][$x]);
    $thisrow++;

}
$listrow = ($productp5[2][0]>3) ? ($listrow + $productp5[2][0]) : ($listrow+3);
//$sheet->setCellValue("L2", $listrow);

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
///* //二、车缝注意事项：*/



/** 三、尺寸注意事项：：*/
$listrow = $listrow +  3 ;

$thisrow = $listrow;
for($x = 0 ,$c = 1; $c <= $productp5[3][0]; $x++ ,$c++){

    if($c >3){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($thisrow, 1);
    }
    $spreadsheet->getActiveSheet()->mergeCells("A{$thisrow}:H{$thisrow}");
    $sheet->setCellValue("A{$thisrow}", $productp5[10][$x]);
    $thisrow++;

}
$listrow = ($productp5[3][0]>3) ? ($listrow + $productp5[3][0]) : ($listrow+3);
//$sheet->setCellValue("L3", $listrow);

$sheet->setCellValue('F'.$listrow, $productp5[3][1]); //处理方法
///* //三、尺寸注意事项：：*/
//
/** 四、洗水注意事项：*/
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

$listrow = $listrow + 1 ;

$thisrow = $listrow;
for($x = 0 ,$c = 1; $c <= $productp5[4][0]; $x++ ,$c++){

    if($c >3){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($thisrow, 1);
    }
    $spreadsheet->getActiveSheet()->mergeCells("A{$thisrow}:H{$thisrow}");
    $sheet->setCellValue("A{$thisrow}", $productp5[11][$x]);
    $thisrow++;

}
$listrow = ($productp5[4][0]>3) ? ($listrow + $productp5[4][0]) : ($listrow+3);

///* 四、洗水注意事项：*/
//
/** 五、整烫注意事项：：*/
$listrow = $listrow +  2 ;

$thisrow = $listrow;
for($x = 0 ,$c = 1; $c <= $productp5[5][0]; $x++ ,$c++){

    if($c >3){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($thisrow, 1);
    }
    $spreadsheet->getActiveSheet()->mergeCells("A{$thisrow}:H{$thisrow}");
    $sheet->setCellValue("A{$thisrow}", $productp5[12][$x]);
    $thisrow++;

}
$listrow = ($productp5[5][0]>3) ? ($listrow + $productp5[5][0]) : ($listrow+3);

///* //五、整烫注意事项：*/
//
/** 六、包装注意事项：*/

$listrow = $listrow +  2 ;

$thisrow = $listrow;
for($x = 0 ,$c = 1; $c <= $productp5[6][0]; $x++ ,$c++){

    if($c >3){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($thisrow, 1);
    }
    $spreadsheet->getActiveSheet()->mergeCells("A{$thisrow}:H{$thisrow}");
    $sheet->setCellValue("A{$thisrow}", $productp5[13][$x]);
    $thisrow++;

}
$listrow = ($productp5[6][0]>3) ? ($listrow + $productp5[6][0]) : ($listrow+3);
//$sheet->setCellValue("L1", $listrow);

if($productp5[6][1] == '1'){
    $radioa = '■ 有';
    $radiob = '□ 无';
}else{
    $radioa = '□ 有';
    $radiob = '■ 无';
}
$sheet->setCellValue('B'.$listrow, $radioa);
$sheet->setCellValue('C'.$listrow, $radiob);

$sheet->setCellValue('F'.$listrow, $productp5[6][2]); //处理方法

/* 六、包装注意事项：*/

/* 七、其他：*/
$listrow = $listrow +  3 ;

$thisrow = $listrow;
for($x = 0 ,$c = 1; $c <= $productp5[7][0]; $x++ ,$c++){

    if($c >3){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($thisrow, 1);
    }
    $spreadsheet->getActiveSheet()->mergeCells("A{$thisrow}:H{$thisrow}");
    $sheet->setCellValue("A{$thisrow}", $productp5[14][$x]);
    $thisrow++;

}
$listrow = ($productp5[7][0]>3) ? ($listrow + $productp5[7][0]) : ($listrow+3);
//$sheet->setCellValue("L1", $listrow);

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