<?php
require_once 'aidenfunc.php';

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器




$fabp1 =   $_SESSION['reportqtyprice'];
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
for($k=0;$k< 11;$k++){
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

$styleArrayRight = [

    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT,
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
//$spreadsheet->getActiveSheet()->setCellValue('E1', 'DATE: '.$fabp1["date"]);
$row = 1;
/**
 * 标题 PHP_EOL
 */
$a = 0;
$titlearr = array('Client','Job no.','Style no.','Qty','FOB price','TOTAL AMOUNT HKD','Actual Delivery','生產外廠','車花價  RMB/PC','工價全包 RMB/PC ');
foreach ($titlearr as $value){
    $col = chr(65 + $a);
    $colname = $col.$row;
    $spreadsheet->getActiveSheet()->getStyle($col.$row)->getAlignment()->setWrapText(true);  //在单元格中写入换行符“\ n”（ALT +“Enter”）
    $spreadsheet->getActiveSheet()->setCellValue($colname, $value);
    $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($styleArray);

    $a++;
}

/**
 * 主体部分
 */
$row++;
if (count($fabp1["cpsid"]) > 0) {

    $clientid_array = $fabp1['clientid'];
    foreach ($clientid_array as $clientid_obj) {

            foreach ($clientid_obj as $client_date) {

                foreach ($client_date as $client_record) {
                    $a = 0;
                    $t = 0;
                    $recordarr = array('clientname','jobno','styleno','qty','fobprice','total','shippingdate','ffname','flower','sewing');

                    foreach ($recordarr as $vc){
                    $col = chr(65 + $t);
                    $record = $recordarr[$t];
                    if($t == 4){
                        $spreadsheet->getActiveSheet()->setCellValue($col.$row, $client_record['fobforex'].$client_record[$record]);
                        $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($styleArray);
                    }elseif($t == 5){
                        $spreadsheet->getActiveSheet()->setCellValue($col.$row, 'HKD '.$client_record[$record]);
                        $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($styleArray);
                    }else{
                        $spreadsheet->getActiveSheet()->setCellValue($col.$row, $client_record[$record]);
                        $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($styleArray);
                    }

                    $t++;
             }
                    $row++;
                }

            }

    }
}
/**
 * 主体部分
 */

/**
 * 总数
 */

$t = 0;
foreach ($titlearr as $value2){
    $col = chr(65 + $t);
    if('Client' == $value2){
        $spreadsheet->getActiveSheet()->setCellValue($col.$row, '总数');
    }
    if('Qty' == $value2){
        $spreadsheet->getActiveSheet()->setCellValue($col.$row, $fabp1["total"]["total"]["totalQty"]);
    }
    if('TOTAL AMOUNT HKD' == $value2){
        $spreadsheet->getActiveSheet()->setCellValue($col.$row, 'HKD '.$fabp1["total"]["total"]["totalHKD"]);
    }
    $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($styleArray);
    $t++;
}


/**
 * 外厂总数
 */
$row++;
$row++;
$row++;
$row++;
/**
 * 标题 PHP_EOL
 */
$a = 0;
$titlearr = array('外厂名称','QTY');
foreach ($titlearr as $value){
    $col = chr(73 + $a);
    $colname = $col.$row;
    $spreadsheet->getActiveSheet()->getStyle($col.$row)->getAlignment()->setWrapText(true);  //在单元格中写入换行符“\ n”（ALT +“Enter”）
    $spreadsheet->getActiveSheet()->setCellValue($colname, $value);
    $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($styleArray);

    $a++;
}


$row++;
if (count($fabp1["cpsid"]) > 0) {
    $ffproduct_array = $fabp1['ffproduct'];

    foreach ($ffproduct_array as $clientid_obj) {

                $a = 0;
                $t = 0;
                $recordarr = array('ffname','qty');

                foreach ($recordarr as $vc){
                    $col = chr(73 + $t);
                    $record = $recordarr[$t];

                    if($t == 1){
                        $spreadsheet->getActiveSheet()->setCellValue($col.$row, $clientid_obj[$record].' pcs');
                        $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($styleArrayRight);
                    }else{
                        $spreadsheet->getActiveSheet()->setCellValue($col.$row, $clientid_obj[$record]);
                        $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($styleArray);
                    }

                    $t++;
                }
                $row++;

    }
    $spreadsheet->getActiveSheet()->setCellValue($col.$row, $fabp1["total"]["total"]["totalQty"].' pcs');
    $spreadsheet->getActiveSheet()->getStyle($col.$row)->applyFromArray($styleArrayRight);
}
/**
 * 外厂总数
 */
$col = 'A';
//range('A','Z')
//$titlearr
foreach (range('A','J') as $item){
    $spreadsheet->getActiveSheet()->getColumnDimension($item)->setAutoSize(true);  //自动列宽度
}

$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页


// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$spreadsheet->setActiveSheetIndex(0);

unset($_SESSION['reportqtyprice'] ); //注销SESSION

$filenameout = "Order Quantity per Month_";
outExcel($spreadsheet,$filenameout);
