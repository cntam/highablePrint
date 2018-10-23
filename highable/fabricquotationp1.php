<?php
session_start();

require_once('autoloadconfig.php');  //判断是否在线

require_once ('img.php');

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器




$fabp1 =   $_SESSION['fabricquotationp1'];
//var_dump($fabp1);

//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/fabricquotationp1.xlsx');
$sheet = $spreadsheet->getActiveSheet();

$spreadsheet->getActiveSheet()->setTitle("第一页");

$spreadsheet->getDefaultStyle()->getFont()->setName('微软雅黑');
$spreadsheet->getDefaultStyle()->getFont()->setSize(12);
//$spreadsheet->getActiveSheet()->getDefaultRowDimension()->setRowHeight(50);


//$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(15);  //列宽度
//$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(10);  //列宽度
//$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(15);  //列宽度

$colarr = range("A","Z");
for($k=0;$k<count($fabp1['title']['a1']);$k++){
    $spreadsheet->getActiveSheet()->getColumnDimension($colarr[$k])->setWidth(15);  //列宽度
}
//$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(15);  //列宽度

$spreadsheet->getActiveSheet()->getPageMargins()->setRight(0.1);
$spreadsheet->getActiveSheet()->getPageMargins()->setLeft(0.1); //页边距

$styleArray1 = [
 'alignment' => [
//        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
//		'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
     'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
     'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP,
    ],
    
//    'borders' => [
//        'top' => [
//            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
//        ],
//
//    ],
   
];


$styleArray = [
    
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


//$spreadsheet->getActiveSheet()->setCellValue('C4', $fabp1["alist"]['a1']);
$spreadsheet->getActiveSheet()->setCellValue('E1', 'DATE: '.$fabp1["date"]);
$row = 6;
/**
 * 标题
 */
$a = 0;
foreach ($fabp1['title']['a1'] as $value){

    $col = chr(65 + $a);
    $colname = $col.$row;
    $spreadsheet->getActiveSheet()->setCellValue($colname, $value);
    $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($styleArray);
    $a++;
}

/**
 * alist
 */
$row++;
for ($y = 0, $i = 1; $i <= $fabp1["alist"]['alistnum']; $i++, $y++) {

    $tdHTML = '';

    for($u = 0,$n = 1;$u< count($fabp1['title']['a1']);$u++,$n++){
        $col = chr(65 + $u);
        if($u == 3){

            $thisvalue = $fabp1["alist"]['a'.$n][$y];
            $n++;
            $issel =  $fabp1["alist"]['a'.$n][$y] == '1' ?  "Y" :  "CM" ;
            $thisvalue .= '/'.$issel;
            $spreadsheet->getActiveSheet()->setCellValue($col.$row, $thisvalue);
        }elseif ($u == 4){
            $thisvalue = $fabp1["alist"]['a'.$n][$y];
            $n++;
            $issel =  $fabp1["alist"]['a'.$n][$y] == '1' ?  "G/M2" :  "G/Y" ;
            $thisvalue .= '/'.$issel;
            $spreadsheet->getActiveSheet()->setCellValue($col.$row, $thisvalue);
        }else{
            $spreadsheet->getActiveSheet()->setCellValue($col.$row, $fabp1["alist"]['a'.$n][$y]);
            $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($styleArray);
        }

    }

    $row++;
}

$row = $row>20 ? $row : 20;
/**
 * remark
 */
$spreadsheet->getActiveSheet()->mergeCells("B{$row}:E{$row}");
$spreadsheet->getActiveSheet()->setCellValue('B'.$row, $fabp1["alist"]['remarks']);
$spreadsheet->getActiveSheet()->getStyle("B{$row}:E{$row}")->applyFromArray($styleArray);
$row++;
foreach ($fabp1["blist"]['b1'] as $value){
    $spreadsheet->getActiveSheet()->mergeCells("B{$row}:E{$row}");
    $spreadsheet->getActiveSheet()->setCellValue('B'.$row, $value);
    $spreadsheet->getActiveSheet()->getStyle("B{$row}:E{$row}")->applyFromArray($styleArray);
    $row++;
}



//
////特殊工序
//if(is_array($fabp1['b10name']) && (count($fabp1['b10name']) >0 )){
//
//    $b10 = '';
//    foreach ($fabp1['b10name'] as $value){
//        $b10 .= '  '.$value;
//    }
//
//}
//
//$spreadsheet->getActiveSheet()->setCellValue('B33', $b10);
//$spreadsheet->getActiveSheet()->setCellValue('B35', $fabp1['data']["blist"]['b11']);
//$spreadsheet->getActiveSheet()->setCellValue('B37', $fabp1['data']["blist"]['b12']);
//
//
//for($i=1,$u=0;$i <= $fabp1['data']["clist"]["clistnum"] ;$i++,$u++){
//    $arownum = 40 + $u;
//
//    $spreadsheet->getActiveSheet()->setCellValue('B'.$arownum, $fabp1['data']["clist"]['c1'][$u]);
//    $spreadsheet->getActiveSheet()->setCellValue('C'.$arownum, $fabp1['data']["clist"]['c2'][$u]);
//    $spreadsheet->getActiveSheet()->setCellValue('D'.$arownum, $fabp1['data']["clist"]['c3'][$u]);
//    $spreadsheet->getActiveSheet()->setCellValue('E'.$arownum, $fabp1['data']["clist"]['c4'][$u]);
//    $spreadsheet->getActiveSheet()->setCellValue('F'.$arownum, $fabp1['data']["clist"]['c5'][$u]);
//    $spreadsheet->getActiveSheet()->setCellValue('G'.$arownum, $fabp1['data']["clist"]['c6'][$u]);
//}
//$spreadsheet->getActiveSheet()->setCellValue('H42', $fabp1['data']["clist"]['c7']);
//
//for($i=1,$u=0;$i <= $fabp1['data']["dlist"]["dlistnum"] ;$i++,$u++){
//    $arownum = 46 + $u;
//
//    $spreadsheet->getActiveSheet()->setCellValue('A'.$arownum, $fabp1['data']["dlist"]['d1'][$u]);
//    $spreadsheet->getActiveSheet()->setCellValue('B'.$arownum, $fabp1['data']["dlist"]['d2'][$u]);
//    $spreadsheet->getActiveSheet()->setCellValue('C'.$arownum, $fabp1['data']["dlist"]['d3'][$u]);
//    $spreadsheet->getActiveSheet()->setCellValue('D'.$arownum, $fabp1['data']["dlist"]['d4'][$u]);
//    $spreadsheet->getActiveSheet()->setCellValue('E'.$arownum, $fabp1['data']["dlist"]['d5'][$u]);
//    $spreadsheet->getActiveSheet()->setCellValue('F'.$arownum, $fabp1['data']["dlist"]['d6'][$u]);
//    $spreadsheet->getActiveSheet()->setCellValue('G'.$arownum, $fabp1['data']["dlist"]['d7'][$u]);
//    $spreadsheet->getActiveSheet()->setCellValue('H'.$arownum, $fabp1['data']["dlist"]['d8'][$u]);
//}
//
////
//
///**
// * 款号备注
// */
//
//$titlearr = array('面布 ：',' 裡布1：',' 裡布2：','裡布3：','撞布1：','撞布2：','撞布3');
//$f = 0;
//$browrow = 25;
//for($r=0;$r<=6;$r++){
//
//    $b1 = 'b'. (1 + $r );
//    $b1 = $fabp1['data']["blist"][$b1];
//    if('on' == $b1){
//        $spreadsheet->getActiveSheet()->setCellValue('B'.$browrow, $titlearr[$r]);
//        $spreadsheet->getActiveSheet()->setCellValue('C'.$browrow, $fabp1['data']["blist"]["b1v"][$r]);
//        $browrow++;
//        $f++;
//    }
//}
//
//if($fabp1['data']["blist"]["formnumb"] > 0){
//    for($r=0 , $k = 1;$k<= $fabp1['data']["blist"]["formnumb"];$r++,$k++){
//
//            $spreadsheet->getActiveSheet()->setCellValue('B'.$browrow, $fabp1['data']["blist"]["b8"][$r]);
//            $spreadsheet->getActiveSheet()->setCellValue('C'.$browrow, $fabp1['data']["blist"]["b9"][$r]);
//        $spreadsheet->getActiveSheet()->getStyle("B{$browrow}")->applyFromArray($styleArray);
//            $browrow++;
//            $f++;
//
//if($f>7){
//    $spreadsheet->getActiveSheet()->insertNewRowBefore($browrow, 1);
//    $spreadsheet->getActiveSheet()->mergeCells("C{$browrow}:H{$browrow}");
//    $spreadsheet->getActiveSheet()->getStyle("B{$browrow}")->applyFromArray($styleArray);
//}
//    }
//}

//
$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页


// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$spreadsheet->setActiveSheetIndex(0);

//unset($_SESSION['samplep1'] ); //注销SESSION

$output=  ($_GET['action'] == 'formdown' )? 1:0;
$nt = date("YmdHis",time()); //转换为日期。
$filenameout = 'fabricquotationp1out'.$nt.'.xlsx';
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
	
    $FILEURL = PRINTURL.$filenameout;
    $MSFILEURL = MSFILEURL. urlencode($FILEURL);

    Header("Location:{$MSFILEURL}");
}
exit;
