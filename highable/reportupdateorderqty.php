<?php
require_once 'aidenfunc.php';

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器




$fabp1 =   $_SESSION['reportupdateorderqty'];
//var_dump($fabp1);


$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$spreadsheet->getActiveSheet()->setTitle("第一页");

$spreadsheet->getDefaultStyle()->getFont()->setName('微软雅黑');
$spreadsheet->getDefaultStyle()->getFont()->setSize(12);
//$spreadsheet->getActiveSheet()->getDefaultRowDimension()->setRowHeight(50);


//$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(15);  //列宽度
//$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(10);  //列宽度
//$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(15);  //列宽度

//$colarr = range("A","Z");
//for($k=0;$k< (count($fabp1['title'])+2);$k++){
//    $spreadsheet->getActiveSheet()->getColumnDimension($colarr[$k])->setWidth(15);  //列宽度
//}
//$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(15);  //列宽度

if(is_array($fabp1['data'])){
    if (count($fabp1['data']) > 0) {
        $col = 'A';
        $colarr =array();
        for($k=0;$k< (count($fabp1['title'])+2);$k++){
            $colarr[] = $col;
            $col++;
        }


        foreach ($colarr as $value){
            $spreadsheet->getActiveSheet()->getColumnDimension($value)->setAutoSize(true);  //自动列宽度
        }



    }
}


$spreadsheet->getActiveSheet()->getPageMargins()->setRight(0.1);
$spreadsheet->getActiveSheet()->getPageMargins()->setLeft(0.1); //页边距




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
$spreadsheet->getActiveSheet()->getStyle('A'.$row)->applyFromArray($bordersLeft);
foreach ($fabp1['title'] as $value){
    $col = chr(66 + $a);
    $colname = $col.$row;
    $spreadsheet->getActiveSheet()->getStyle($col.$row)->getAlignment()->setWrapText(true);  //在单元格中写入换行符“\ n”（ALT +“Enter”）
    $spreadsheet->getActiveSheet()->setCellValue($colname, $value);
    $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($bordersLeft);

    $a++;
}
$col = chr(66 + $a);
$spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($bordersLeft);
/**
 * 主体部分
 */
$row++;
if (count($fabp1['data']) > 0) {

    $clientid_array = $fabp1['data'];
    foreach ($clientid_array as $obj_key=>$obj_value) {

        $month = monthStr($obj_key);

        $spreadsheet->getActiveSheet()->setCellValue('A'.$row, $month);
        $spreadsheet->getActiveSheet()->getStyle('A'.$row)->applyFromArray($bordersLeft);
        $t = 0;
        foreach ($obj_value as $client_record){
            $col = chr(66 + $t);
            if(is_array($client_record)){
                $spreadsheet->getActiveSheet()->setCellValue($col.$row, $client_record["qty"]);
            };
            $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($bordersLeft);
            $t++;

            }
        $col = chr(66 + $t);
        $spreadsheet->getActiveSheet()->setCellValue($col.$row, $fabp1['total']['monthqty'][$obj_key]);
        $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($bordersLeft);
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
    $spreadsheet->getActiveSheet()->getStyle('A'.$row)->applyFromArray($bordersLeft);

    $clientid_array = $fabp1['total']["singletotalqty"];
    foreach ($clientid_array as $obj_key=>$obj_value) {
        $col = chr(66 + $t);
        $spreadsheet->getActiveSheet()->setCellValue($col.$row, $obj_value);
        $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($bordersLeft);
        $t++;
    }
    $col = chr(66 + $t);
    $spreadsheet->getActiveSheet()->setCellValue($col.$row, $fabp1["total"]["totalAllQty"]);
    $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($bordersLeft);

    $row++;
}
if (count($fabp1['total']["Qtypercent"]) > 0) {
    $t = 0;
    $spreadsheet->getActiveSheet()->getStyle('A'.$row)->applyFromArray($bordersLeft);
    $clientid_array = $fabp1['total']["Qtypercent"];
    foreach ($clientid_array as $obj_key=>$obj_value) {
        $col = chr(66 + $t);
        $spreadsheet->getActiveSheet()->setCellValue($col.$row, $obj_value);
        $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($bordersLeft);
        $sheet->getStyle($col.$row)->getNumberFormat()->setFormatCode('0.00%');
        //$spreadsheet->getActiveSheet()->getStyle($col.$row)->getNumberFormat()->applyFromArray( [ 'formatCode' => NumberFormat::FORMAT_PERCENTAGE_00 ] );
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
$spreadsheet->getActiveSheet()->getStyle('A'.$row)->applyFromArray($bordersLeft);
foreach ($fabp1['title'] as $value){
    $col = chr(66 + $a);
    $colname = $col.$row;
    $spreadsheet->getActiveSheet()->getStyle($col.$row)->getAlignment()->setWrapText(true);  //在单元格中写入换行符“\ n”（ALT +“Enter”）
    $spreadsheet->getActiveSheet()->setCellValue($colname, $value);
    $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($bordersLeft);

    $a++;
}
$col = chr(66 + $a);
$spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($bordersLeft);
/**
 * 主体部分
 */
$row++;
if(is_array($fabp1['data'])){
    if (count($fabp1['data']) > 0) {

        $clientid_array = $fabp1['data'];
        foreach ($clientid_array as $obj_key=>$obj_value) {

            $month = monthStr($obj_key);

            $spreadsheet->getActiveSheet()->setCellValue('A'.$row, $month);
            $spreadsheet->getActiveSheet()->getStyle('A'.$row)->applyFromArray($bordersLeft);
            $t = 0;
            foreach ($obj_value as $client_record){
                $col = chr(66 + $t);
                if(is_array($client_record)){
                    $spreadsheet->getActiveSheet()->setCellValue($col.$row, $client_record["HKD"]);
                };
                $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($bordersLeft);
                $sheet->getStyle($col.$row)->getNumberFormat()->setFormatCode('"HK$"#,##0.00_-');
                $t++;

            }
            $col = chr(66 + $t);
            $spreadsheet->getActiveSheet()->setCellValue($col.$row, $fabp1['total']['monthHKD'][$obj_key]);
            $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($bordersLeft);
            $sheet->getStyle($col.$row)->getNumberFormat()->setFormatCode('"HK$"#,##0.00_-');
            $row++;
        }
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
    $spreadsheet->getActiveSheet()->getStyle('A'.$row)->applyFromArray($bordersLeft);

    $clientid_array = $fabp1['total']["singletotalHKD"];
    foreach ($clientid_array as $obj_key=>$obj_value) {
        $col = chr(66 + $t);
        $spreadsheet->getActiveSheet()->setCellValue($col.$row, $obj_value);
        $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($bordersLeft);
        $sheet->getStyle($col.$row)->getNumberFormat()->setFormatCode('"HK$"#,##0.00_-');
        $t++;
    }
    $col = chr(66 + $t);
    $spreadsheet->getActiveSheet()->setCellValue($col.$row, $fabp1["total"]["totalAllHKD"]);
    $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($bordersLeft);
    $sheet->getStyle($col.$row)->getNumberFormat()->setFormatCode('"HK$"#,##0.00_-');
    $row++;
}

//$sheet->getStyle('A1')
//    ->getNumberFormat()
//    ->setFormatCode('"HK$"#,##0.00_-');


$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页



// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$spreadsheet->setActiveSheetIndex(0);

unset($_SESSION['reportupdateorderqty'] ); //注销SESSION


$filenameout = "Order Quantity per Year";
outExcel($spreadsheet,$filenameout);

