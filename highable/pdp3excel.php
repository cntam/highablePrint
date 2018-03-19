<?php
session_start();

$client =  $_SESSION['client'] ;
$date =    $_SESSION['date'] ;
$remark1 = $_SESSION['remark1'] ;
$remark2 = $_SESSION['remark2'] ;
$remark3 = $_SESSION['remark3'] ;
$remark4 = $_SESSION['remark4'] ;
$remark5 = $_SESSION['remark5'] ;
$remark6 = $_SESSION['remark6'] ;
$remark7 = $_SESSION['remark7'] ;
$remark8 = $_SESSION['remark8'] ;


require '/home/pan/vendor/autoload.php';
//require '../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

$spreadsheet->getActiveSheet()->setTitle("sheet1");
//$sheet->setCellValue('A1', 'Hello World !');
$spreadsheet->getDefaultStyle()->getFont()->setName('微软雅黑');
$spreadsheet->getDefaultStyle()->getFont()->setSize(12);
$spreadsheet->getActiveSheet()->getDefaultRowDimension()->setRowHeight(50);
$styleArray1 = [
 'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
		'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
    ],
    
    'borders' => [
        'top' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
        ],
		
    ],
   
];
$spreadsheet->getActiveSheet()->getStyle('A1:D1')->applyFromArray($styleArray1);

	$spreadsheet->getActiveSheet()->getStyle('A1')
    ->getBorders()->getLEFT()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK);
		$spreadsheet->getActiveSheet()->getStyle('A1')
    ->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
			$spreadsheet->getActiveSheet()->getStyle('B1')
    ->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);

$spreadsheet->getActiveSheet()->getStyle('D1')
    ->getBorders()->getLEFT()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
	$spreadsheet->getActiveSheet()->getStyle('D1')
    ->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK);



$styleArray = [
    
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
		'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP,
    ],
	
    'borders' => [
        'top' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
        ],
		'bottom' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
        ],
		'left' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
        ],
		'right' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
        ],
    ],
   
];

$spreadsheet->getActiveSheet()->getStyle('A2:B2')->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle('C2:D2')->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle('A3:B3')->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle('C3:D3')->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle('A4:B4')->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle('C4:D4')->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle('A5:B5')->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle('C5:D5')->applyFromArray($styleArray);

$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(15);
$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(19);
$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(15);
$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(19);

$spreadsheet->getActiveSheet()->getRowDimension('1')->setRowHeight(40);
$spreadsheet->getActiveSheet()->getRowDimension('2')->setRowHeight(160);
$spreadsheet->getActiveSheet()->getRowDimension('3')->setRowHeight(160);
$spreadsheet->getActiveSheet()->getRowDimension('4')->setRowHeight(160);
$spreadsheet->getActiveSheet()->getRowDimension('5')->setRowHeight(160);


// Set cell A1 with a string value

//$spreadsheet->getActiveSheet()->getStyle('A1:D1')->getFont()->setSize(14);

$spreadsheet->getActiveSheet()->setCellValue('A1', 'CLIENT:');
$spreadsheet->getActiveSheet()->setCellValue('B1', "$client");
$spreadsheet->getActiveSheet()->setCellValue('C1', "DATE:");
$spreadsheet->getActiveSheet()->setCellValue('D1', "$date");

// Set cell A2 with a numeric value.

//$spreadsheet->getActiveSheet()->getColumnDimension('A')->setAutoSize(true);
$spreadsheet->getActiveSheet()->mergeCells('A2:B2');
$spreadsheet->getActiveSheet()->setCellValue('A2', "$remark1");
$spreadsheet->getActiveSheet()->mergeCells('C2:D2');
$spreadsheet->getActiveSheet()->setCellValue('C2', "$remark2");

// Set cell A2 with a numeric value
$spreadsheet->getActiveSheet()->mergeCells('A3:B3');
$spreadsheet->getActiveSheet()->setCellValue('A3', "$remark3");
$spreadsheet->getActiveSheet()->mergeCells('C3:D3');
$spreadsheet->getActiveSheet()->setCellValue('C3', "$remark4");

// Set cell A2 with a numeric value
$spreadsheet->getActiveSheet()->mergeCells('A4:B4');
$spreadsheet->getActiveSheet()->setCellValue('A4', "$remark5");
$spreadsheet->getActiveSheet()->mergeCells('C4:D4');
$spreadsheet->getActiveSheet()->setCellValue('C4', "$remark6");

// Set cell A2 with a numeric value
$spreadsheet->getActiveSheet()->mergeCells('A5:B5');
$spreadsheet->getActiveSheet()->setCellValue('A5', "$remark7");
$spreadsheet->getActiveSheet()->mergeCells('C5:D5');
$spreadsheet->getActiveSheet()->setCellValue('C5', "$remark8");

$spreadsheet->getActiveSheet()->getStyle('A2:C2')->getAlignment()->setWrapText(true);
$spreadsheet->getActiveSheet()->getStyle('A3:C3')->getAlignment()->setWrapText(true);
$spreadsheet->getActiveSheet()->getStyle('A4:C4')->getAlignment()->setWrapText(true);
$spreadsheet->getActiveSheet()->getStyle('A5:C5')->getAlignment()->setWrapText(true);



// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$spreadsheet->setActiveSheetIndex(0);

$output=  ($_GET['action'] == 'formdown' )? 1:0;
$filenameout = 'pdp3out.xlsx';
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
    //echo "<a href= 'http://view.officeapps.live.com/op/view.aspx?src=". urlencode($FILEURL)."' target='_blank' >跳轉--{$filename}</a>";
    Header("Location:{$MSFILEURL}");

    //Header("Location:{$printURL}");
}
exit;
