<?php
session_start();

$costp2 =  $_SESSION['costp2'];

require '/home/pan/vendor/autoload.php';
//require '/home/soft/vendor/autoload.php';
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
$spreadsheet->getDefaultStyle()->getFont()->setSize(10);
//$spreadsheet->getActiveSheet()->getDefaultRowDimension()->setRowHeight(50); //行默认高度

$styleArray1 = [
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
        'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
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



//合并/取消合并单元格
$spreadsheet->getActiveSheet()->mergeCells("B1:C1");
$spreadsheet->getActiveSheet()->mergeCells("B2:C2");
$spreadsheet->getActiveSheet()->mergeCells("B3:C3");
$spreadsheet->getActiveSheet()->mergeCells("B4:C11");
for ($i = 12;$i<41;$i++){
    $spreadsheet->getActiveSheet()->mergeCells("B{$i}:C{$i}");
    $spreadsheet->getActiveSheet()->getStyle("B{$i}:C{$i}")->applyFromArray($styleArray1);
}
$spreadsheet->getActiveSheet()->mergeCells("A4:A11");
//$spreadsheet->getActiveSheet()->mergeCells("B4:C11");



/*
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

*/

$styleArray = [
    
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
		'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
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
for ($i = 1;$i<44;$i++) {
    $spreadsheet->getActiveSheet()->getStyle("A{$i}")->applyFromArray($styleArray);
}


$styleArray2 = [

    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
        'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
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
for ($i = 1;$i<44;$i++) {

    $spreadsheet->getActiveSheet()->getStyle("B{$i}")->applyFromArray($styleArray2);
    $spreadsheet->getActiveSheet()->getStyle("C{$i}")->applyFromArray($styleArray2);
}
/*
$spreadsheet->getActiveSheet()->getStyle('C2:D2')->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle('A3:B3')->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle('C3:D3')->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle('A4:B4')->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle('C4:D4')->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle('A5:B5')->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle('C5:D5')->applyFromArray($styleArray);
*/
$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(32); //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(21);
$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(32);
/*


$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(19);

$spreadsheet->getActiveSheet()->getRowDimension('1')->setRowHeight(40);
$spreadsheet->getActiveSheet()->getRowDimension('2')->setRowHeight(160);
$spreadsheet->getActiveSheet()->getRowDimension('3')->setRowHeight(160);
$spreadsheet->getActiveSheet()->getRowDimension('4')->setRowHeight(160);
$spreadsheet->getActiveSheet()->getRowDimension('5')->setRowHeight(160);
*/

// Set cell A1 with a string value

//$spreadsheet->getActiveSheet()->getStyle('A1:D1')->getFont()->setSize(14);


$spreadsheet->getActiveSheet()->setCellValue('B1', $costp2['costname']);
$spreadsheet->getActiveSheet()->setCellValue('A2', "DATE:");
$spreadsheet->getActiveSheet()->setCellValue('B2', $costp2['costdata']);
$spreadsheet->getActiveSheet()->setCellValue('A3', "款式：");
$spreadsheet->getActiveSheet()->setCellValue('B3', $costp2['costno']);


//$img = 'http://www.a.cn/wordpress/wp-content/uploads/2018/02/2506390415_532601864.220x220-20.jpg';
$img = $costp2['remarkimg2'];
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

//$img = imagecreatefromjpeg($img);

$width = imagesx($img);

$height = imagesy($img);


// Generate an image
//$gdImage = @imagecreatetruecolor($width, $height) or die('Cannot Initialize new GD image stream');
//$textColor = imagecolorallocate($gdImage, 255, 255, 255);
//imagestring($gdImage, 1, 5, 5,  'Created with PhpSpreadsheet', $textColor);

// Add a drawing to the worksheet
$drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
$drawing->setName($costp2['costname']);
$drawing->setDescription($costp2['costname']);
//$drawing->setImageResource($gdImage);
$drawing->setImageResource($img);
$drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
$drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
//$drawing->setHeight($width);

//$drawing->setHeight($width>550 ? 550:$width);
//$drawing->setWidth(300);
$drawing->setHeight(135);
$drawing->setCoordinates('B4');
$drawing->setOffsetX(130);
$drawing->setOffsetY(5);
$drawing->setWorksheet($spreadsheet->getActiveSheet());



$spreadsheet->getActiveSheet()->setCellValue('A12', "Style no：");
$spreadsheet->getActiveSheet()->setCellValue('B12', $costp2['styleno']);


$fabname=array("SLEF FABRIC：","Fabric Cost：","Cons./Doz(NET)：",'CONTRAST 1 fabric：','Fabric Cost：','Cons./Doz(NET)：','CONTRAST 2 fabric：','Fabric Cost：','Cons./Doz(NET)：','CONTRAST 3 fabric：','Fabric Cost：','Cons./Doz(NET)：','FABRIC COST：','Interlining @ 15：','Thread：','MCQ label(main label,size label & CO label)：','Carton：','MCQ poly bag & hangtag：','Fabric test cost：','Sticker：','18L Shell button(7+1) use for centre front placker：','16L Shell button(2+1) use for cuff：','trimming cost：','Tatal trim cost(10%)：','Sewing(RMB:120.0/PC)：','Cut,Trim,Pack etc.','Factory Overhead','Profit margin','100-200PCS+CM30%','201-400PCS+CM15%','OVER 400PCS');
$fabcou = count($fabname);
for ($j=0,$k = 13,$v =1 ;$j<$fabcou;$j++){

    $m = $k+2;
    if($v<29){
        $spreadsheet->getActiveSheet()->setCellValue("A{$k}", $fabname[$j]);
        $spreadsheet->getActiveSheet()->setCellValue("B{$k}", $costp2['fab']['a'.$v]);
    }else{
        if($v == 29){
            $spreadsheet->getActiveSheet()->mergeCells("A{$k}:A{$m}");
            $spreadsheet->getActiveSheet()->setCellValue("A{$k}", 'Unit Price');
            $spreadsheet->getActiveSheet()->setCellValue("B{$k}", $fabname[$j]);
            $spreadsheet->getActiveSheet()->setCellValue("C{$k}", $costp2['fab']['a'.$v]);

        }else{
            $spreadsheet->getActiveSheet()->setCellValue("B{$k}", $fabname[$j]);
            $spreadsheet->getActiveSheet()->setCellValue("C{$k}", $costp2['fab']['a'.$v]);
        }


    }
    $spreadsheet->getActiveSheet()->getStyle("A{$k}")->getAlignment()->setWrapText(true);

    $k++;
    $v++;
}

// Set cell A2 with a numeric value.
/*
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
*/


// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$spreadsheet->setActiveSheetIndex(0);

unset($_SESSION['costp2'] ); //注销SESSION

$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页
$output=  ($_GET['action'] == 'formdown' )? 1:0;
$nt = date("YmdHis",time()); //转换为日期。
$filenameout = 'costp2out'.$nt.'.xlsx';
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