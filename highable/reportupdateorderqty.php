<?php
session_start();

require_once('autoloadconfig.php');  //判断是否在线

require_once ('img.php');

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器




$fabp1 =   $_SESSION['reportupdateorderqty'];
//var_dump($fabp1);

//$spreadsheet = new Spreadsheet();
//$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/fabricquotationp1.xlsx');
//$sheet = $spreadsheet->getActiveSheet();

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$spreadsheet->getActiveSheet()->setTitle("第一页");

$spreadsheet->getDefaultStyle()->getFont()->setName('微软雅黑');
$spreadsheet->getDefaultStyle()->getFont()->setSize(12);
//$spreadsheet->getActiveSheet()->getDefaultRowDimension()->setRowHeight(50);


//$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(15);  //列宽度
//$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(10);  //列宽度
//$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(15);  //列宽度

$colarr = range("A","Z");
for($k=0;$k< (count($fabp1['title'])+2);$k++){
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
function monthStr($in){
    $month = substr($in,4);

    switch ($month)
    {
        case '01':
            $output = "JAN";
            break;
        case '02':
            $output = "FEB";
            break;
        case '03':
            $output = "MAR";
            break;
        case '04':
            $output = "APRIL";
            break;
        case '05':
            $output = "MAY";
            break;
        case '06':
            $output = "JUNE";
            break;
        case '07':
            $output = "JULY";
            break;
        case '08':
            $output = "AUG";
            break;
        case '09':
            $output = "SEPT";
            break;
        case '10':
            $output = "OCT";
            break;
        case '11':
            $output = "NOV";
            break;
        case '12':
            $output = "DEC";
            break;
        default:
            $output = "";
    }

    return $output;
}

//$spreadsheet->getActiveSheet()->setCellValue('C4', $fabp1["alist"]['a1']);
//$spreadsheet->getActiveSheet()->setCellValue('E1', 'DATE: '.$fabp1["date"]);
$row = 1;
/**
 * 标题 PHP_EOL
 */
$a = 0;
$spreadsheet->getActiveSheet()->getStyle('A'.$row)->applyFromArray($styleArray);
foreach ($fabp1['title'] as $value){
    $col = chr(66 + $a);
    $colname = $col.$row;
    $spreadsheet->getActiveSheet()->getStyle($col.$row)->getAlignment()->setWrapText(true);  //在单元格中写入换行符“\ n”（ALT +“Enter”）
    $spreadsheet->getActiveSheet()->setCellValue($colname, $value);
    $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($styleArray);

    $a++;
}
$col = chr(66 + $a);
$spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($styleArray);
/**
 * 主体部分
 */
$row++;
if (count($fabp1['data']) > 0) {

    $clientid_array = $fabp1['data'];
    foreach ($clientid_array as $obj_key=>$obj_value) {

        $month = monthStr($obj_key);

        $spreadsheet->getActiveSheet()->setCellValue('A'.$row, $month);
        $spreadsheet->getActiveSheet()->getStyle('A'.$row)->applyFromArray($styleArray);
        $t = 0;
        foreach ($obj_value as $client_record){
            $col = chr(66 + $t);
            if(is_array($client_record)){
                $spreadsheet->getActiveSheet()->setCellValue($col.$row, $client_record["qty"]);
            };
            $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($styleArray);
            $t++;

            }
        $col = chr(66 + $t);
        $spreadsheet->getActiveSheet()->setCellValue($col.$row, $fabp1['total']['monthqty'][$obj_key]);
        $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($styleArray);
        $row++;
    }
}

/**
 * 主体部分
 */

/**
 * 总数
 */
if (count($fabp1['total']) > 0) {
    $t = 0;
    $spreadsheet->getActiveSheet()->setCellValue('A'.$row, '总数');
    $spreadsheet->getActiveSheet()->getStyle('A'.$row)->applyFromArray($styleArray);

    $clientid_array = $fabp1['total']["singletotalqty"];
    foreach ($clientid_array as $obj_key=>$obj_value) {
        $col = chr(66 + $t);
        $spreadsheet->getActiveSheet()->setCellValue($col.$row, $obj_value);
        $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($styleArray);
        $t++;
    }
    $col = chr(66 + $t);
    $spreadsheet->getActiveSheet()->setCellValue($col.$row, $fabp1["total"]["totalAllQty"]);
    $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($styleArray);

    $row++;
}
if (count($fabp1['total']["Qtypercent"]) > 0) {
    $t = 0;
    $spreadsheet->getActiveSheet()->getStyle('A'.$row)->applyFromArray($styleArray);
    $clientid_array = $fabp1['total']["Qtypercent"];
    foreach ($clientid_array as $obj_key=>$obj_value) {
        $col = chr(66 + $t);
        $spreadsheet->getActiveSheet()->setCellValue($col.$row, $obj_value);
        $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($styleArray);
        $t++;
    }
    $row++;
}
$row++;
$row++;

/**
 * 标题 PHP_EOL
 */
$a = 0;
$spreadsheet->getActiveSheet()->getStyle('A'.$row)->applyFromArray($styleArray);
foreach ($fabp1['title'] as $value){
    $col = chr(66 + $a);
    $colname = $col.$row;
    $spreadsheet->getActiveSheet()->getStyle($col.$row)->getAlignment()->setWrapText(true);  //在单元格中写入换行符“\ n”（ALT +“Enter”）
    $spreadsheet->getActiveSheet()->setCellValue($colname, $value);
    $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($styleArray);

    $a++;
}
$col = chr(66 + $a);
$spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($styleArray);
/**
 * 主体部分
 */
$row++;
if (count($fabp1['data']) > 0) {

    $clientid_array = $fabp1['data'];
    foreach ($clientid_array as $obj_key=>$obj_value) {

        $month = monthStr($obj_key);

        $spreadsheet->getActiveSheet()->setCellValue('A'.$row, $month);
        $spreadsheet->getActiveSheet()->getStyle('A'.$row)->applyFromArray($styleArray);
        $t = 0;
        foreach ($obj_value as $client_record){
            $col = chr(66 + $t);
            if(is_array($client_record)){
                $spreadsheet->getActiveSheet()->setCellValue($col.$row, $client_record["HKD"]);
            };
            $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($styleArray);
            $t++;

        }
        $col = chr(66 + $t);
        $spreadsheet->getActiveSheet()->setCellValue($col.$row, $fabp1['total']['monthHKD'][$obj_key]);
        $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($styleArray);
        $row++;
    }
}

/**
 * 主体部分
 */

/**
 * 总数
 */
if (count($fabp1['total']) > 0) {
    $t = 0;
    $spreadsheet->getActiveSheet()->setCellValue('A'.$row, '总数');
    $spreadsheet->getActiveSheet()->getStyle('A'.$row)->applyFromArray($styleArray);

    $clientid_array = $fabp1['total']["singletotalHKD"];
    foreach ($clientid_array as $obj_key=>$obj_value) {
        $col = chr(66 + $t);
        $spreadsheet->getActiveSheet()->setCellValue($col.$row, $obj_value);
        $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($styleArray);
        $t++;
    }
    $col = chr(66 + $t);
    $spreadsheet->getActiveSheet()->setCellValue($col.$row, $fabp1["total"]["totalAllHKD"]);
    $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($styleArray);

    $row++;
}




$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页


// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$spreadsheet->setActiveSheetIndex(0);

//unset($_SESSION['samplep1'] ); //注销SESSION

$output=  ($_GET['action'] == 'formdown' )? 1:0;
$nt = date("YmdHis",time()); //转换为日期。
$filenameout = 'reportupdateorderqtyout'.$nt.'.xlsx';
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
