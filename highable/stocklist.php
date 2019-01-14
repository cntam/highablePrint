<?php
session_start();

require_once('autoloadconfig.php');  //判断是否在线

require_once ('img.php');

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器


$fabp1 =   $_SESSION['stocklist'];
//var_dump($fabp1);

$spreadsheet = new Spreadsheet();
//$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/fabricquotationp1.xlsx');
$sheet = $spreadsheet->getActiveSheet();

$sheet->setTitle("第一页");

$spreadsheet->getDefaultStyle()->getFont()->setName('微软雅黑');
$spreadsheet->getDefaultStyle()->getFont()->setSize(12);
$sheet->getColumnDimension('F')->setWidth(18);  //列宽度 修改

$colarr = range("A","Z");
for($k=0;$k<count($fabp1['title']);$k++){
    $sheet->getColumnDimension($colarr[$k])->setWidth(15);  //列宽度
}


$sheet->getPageMargins()->setRight(0.1);
$sheet->getPageMargins()->setLeft(0.1); //页边距

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
        'wrapText' => true,
        'ShrinkToFit'=>true,
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

$styleArray2 = [

    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
        'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP,
        'wrapText' => true,
        'ShrinkToFit'=>true,
    ],
    'font' => [
        'bold' => true,
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

//填数据 修改
$sheet->mergeCells("A1:C1");

if ($fabp1['action'] == 'material') {
    $sheet->setCellValue('A1', 'Material_Stock_List');
} else {
    $sheet->setCellValue('A1', 'Fabric_Stock_List');
}


$sheet->setCellValue('A3', '客户');
$sheet->getStyle('A3')->applyFromArray($styleArray2);

$row = 3;
/**
 * 标题
 */
$a = 0;
foreach ($fabp1['title'] as $value){

    $col = chr(66 + $a); // B
    $colname = $col.$row;
    $sheet->setCellValue($colname, $value);
    $sheet->getStyle($col.$row)->applyFromArray($styleArray2);
    $a++;
}

/**
 * content
 */
$a = 0;
foreach ($fabp1["info"] as $value) {
    $row = 4;
    $col = chr(65 + $a); // A
//    $colname = $col.$row;
    foreach ($value as $childvalue) {
        $sheet->getStyle($col.$row)->applyFromArray($styleArray);
        $sheet->setCellValue($col.$row, $childvalue);
        $row++;
    }
    $a++;
}



$sheet->getPageSetup()->setFitToPage(true); //将工作表调整为一页


// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$spreadsheet->setActiveSheetIndex(0);

unset($_SESSION['stocklist'] ); //注销SESSION

$output=  ($_GET['action'] == 'formdown' )? 1:0;
$nt = date("mdHis",time()); //转换为日期。
// 修改 根据不同action
if ($fabp1['action'] == 'material') {
    $filenameout = 'Material_Stock_List'.$nt.'.xlsx';
} else if ($fabp1['action'] == 'fabric') {
    $filenameout = 'Fabric_Stock_List'.$nt.'.xlsx';
}

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
}
exit;
