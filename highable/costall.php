<?php
session_start();

$costall =  $_SESSION['costall'];
require '/home/pan/vendor/autoload.php';
//require '/home/soft/vendor/autoload.php';
//require '../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;


    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/costp1.xlsx');

    $sheet = $spreadsheet->getActiveSheet();


    $spreadsheet->getActiveSheet()->setTitle("sheet1");
//$sheet->setCellValue('A1', 'Hello World !');
    $spreadsheet->getDefaultStyle()->getFont()->setName('微软雅黑');
    $spreadsheet->getDefaultStyle()->getFont()->setSize(12);
//$spreadsheet->getActiveSheet()->getDefaultRowDimension()->setRowHeight(50);  //默认行高度

    $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(5);//列宽度高度
    $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(13);//列宽度高度
    $spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(5);//列宽度高度

//$spreadsheet->getActiveSheet()->setCellValue('B1', 'IHK NO.');

    $img = $costall['costp1']['remarkimg2'];

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
    $drawing->setName($costall["costp1"]['costname']);
    $drawing->setDescription($costall["costp1"]['costname']);
//$drawing->setImageResource($gdImage);
    $drawing->setImageResource($img);
    $drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
    $drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
//$drawing->setHeight($width);

//$drawing->setHeight($width>550 ? 550:$width);
//$drawing->setWidth($width>500 ? 500:$width);

$resw = $width < 660 ? 0 : 2;
$resh = $height < 650 ? 0 : 3;
$res = $resw + $resh;
switch ($res)
{
    case "2":
        $drawing->setWidth(660);
        break;
    case "3":
        $drawing->setHeight(650);
        break;
    case "5":
        $drawing->setWidth(660);
        break;

    default:
        $drawing->setHeight($height>650 ? 650:$height);
}


    $drawing->setCoordinates('A1');
    $drawing->setOffsetX(10);
    $drawing->setOffsetY(10);
    $drawing->setWorksheet($spreadsheet->getActiveSheet());


    $styleArray1 = [
        'font' => [
            'bold' => true,
        ],
        'alignment' => [
            'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
            'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
        ],

    ];
    /*$spreadsheet->getActiveSheet()->getStyle('B3')->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->getStyle('B4')->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->getStyle('B5')->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->getStyle('B6')->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->getStyle('B7')->applyFromArray($styleArray1);*/



    $styleArray = [

        'alignment' => [
            'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
            'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
        ],
        //$spreadsheet->getActiveSheet()->getStyle('A2:C2')->getAlignment()->setShrinkToFit(true);//缩小以适合
        'borders' => [

            'bottom' => [
                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
            ],

        ],

    ];

    /*$spreadsheet->getActiveSheet()->getStyle('D3:G3')->applyFromArray($styleArray);
    $spreadsheet->getActiveSheet()->getStyle('D4:G4')->applyFromArray($styleArray);
    $spreadsheet->getActiveSheet()->getStyle('D5:G5')->applyFromArray($styleArray);
    $spreadsheet->getActiveSheet()->getStyle('D6:G6')->applyFromArray($styleArray);
    $spreadsheet->getActiveSheet()->getStyle('D7:G7')->applyFromArray($styleArray);*/



//$spreadsheet->getActiveSheet()->getRowDimension('3')->setRowHeight(30); //行高度



// Set cell A1 with a string value

//$spreadsheet->getActiveSheet()->getStyle('A1:D1')->getFont()->setSize(14);

//填数据

    $spreadsheet->getActiveSheet()->setCellValue('B35', $costall['costp1']['costname']);
    $spreadsheet->getActiveSheet()->setCellValue('A36', $costall['costno']);
    $spreadsheet->getActiveSheet()->setCellValue('B36', $costall['costp1']['ccno2']);
    $spreadsheet->getActiveSheet()->setCellValue('E38', $costall['costno']);

    $spreadsheet->getActiveSheet()->setCellValue('B38', $costall['costp1']['fab']['a1']);
    $spreadsheet->getActiveSheet()->setCellValue('B39', $costall['costp1']['fab']['b1']);
    $spreadsheet->getActiveSheet()->setCellValue('B40', $costall['costp1']['fab']['c1']);
    $spreadsheet->getActiveSheet()->setCellValue('B41', $costall['costp1']['fab']['d1']);
    $spreadsheet->getActiveSheet()->setCellValue('B42', $costall['costp1']['fab']['e1']);
    $spreadsheet->getActiveSheet()->setCellValue('B43', $costall['costp1']['fab']['f1']);
    $spreadsheet->getActiveSheet()->setCellValue('B44', $costall['costp1']['fab']['g1']);
    $spreadsheet->getActiveSheet()->setCellValue('B45', $costall['costp1']['fab']['h1']);

    $spreadsheet->getActiveSheet()->setCellValue('E40', $costall['costp1']['fab']['c2']);
    $spreadsheet->getActiveSheet()->setCellValue('E41', $costall['costp1']['fab']['d2']);
    $spreadsheet->getActiveSheet()->setCellValue('E42', $costall['costp1']['fab']['e2']);
    $spreadsheet->getActiveSheet()->setCellValue('E43', $costall['costp1']['fab']['f2']);
    $spreadsheet->getActiveSheet()->setCellValue('E44', $costall['costp1']['fab']['g2']);
    $spreadsheet->getActiveSheet()->setCellValue('E45', $costall['costp1']['fab']['h2']);

    $spreadsheet->getActiveSheet()->setCellValue('H40', $costall['costp1']['fab']['c3']);
    $spreadsheet->getActiveSheet()->setCellValue('H41', $costall['costp1']['fab']['d3']);
    $spreadsheet->getActiveSheet()->setCellValue('H42', $costall['costp1']['fab']['e3']);
    $spreadsheet->getActiveSheet()->setCellValue('H43', $costall['costp1']['fab']['f3']);
    $spreadsheet->getActiveSheet()->setCellValue('H44', $costall['costp1']['fab']['g3']);


    /**
    第二頁
     */
$spreadsheet->setActiveSheetIndex(1);  //設置當前活動表
$spreadsheet->getActiveSheet()->setTitle("sheet2");
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


$spreadsheet->getActiveSheet()->setCellValue('B1', $costall['costp2']['costname']);
$spreadsheet->getActiveSheet()->setCellValue('A2', "DATE:");
$spreadsheet->getActiveSheet()->setCellValue('B2', $costall['costp2']['costdata']);
$spreadsheet->getActiveSheet()->setCellValue('A3', "款式：");
$spreadsheet->getActiveSheet()->setCellValue('B3', $costall['costno']);


//$img = 'http://www.a.cn/wordpress/wp-content/uploads/2018/02/2506390415_532601864.220x220-20.jpg';
$img = $costall['costp2']['remarkimg2'];

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
$drawing->setName($costall['costp2']['costname']);
$drawing->setDescription($costall['costp2']['costname']);
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
$spreadsheet->getActiveSheet()->setCellValue('B12', $costall['costp2']['styleno']);


$fabname=array("SLEF FABRIC：","Fabric Cost：","Cons./Doz(NET)：",'CONTRAST 1 fabric：','Fabric Cost：','Cons./Doz(NET)：','CONTRAST 2 fabric：','Fabric Cost：','Cons./Doz(NET)：','CONTRAST 3 fabric：','Fabric Cost：','Cons./Doz(NET)：','FABRIC COST：','Interlining @ 15：','Thread：','MCQ label(main label,size label & CO label)：','Carton：','MCQ poly bag & hangtag：','Fabric test cost：','Sticker：','18L Shell button(7+1) use for centre front placker：','16L Shell button(2+1) use for cuff：','trimming cost：','Tatal trim cost(10%)：','Sewing(RMB:120.0/PC)：','Cut,Trim,Pack etc.','Factory Overhead','Profit margin','100-200PCS+CM30%','201-400PCS+CM15%','OVER 400PCS');
$fabcou = count($fabname);
for ($j=0,$k = 13,$v =1 ;$j<$fabcou;$j++){

    $m = $k+2;
    if($v<29){
        $spreadsheet->getActiveSheet()->setCellValue("A{$k}", $fabname[$j]);
        $spreadsheet->getActiveSheet()->setCellValue("B{$k}", $costall['costp2']['fab']['a'.$v]);
    }else{
        if($v == 29){
            $spreadsheet->getActiveSheet()->mergeCells("A{$k}:A{$m}");
            $spreadsheet->getActiveSheet()->setCellValue("A{$k}", 'Unit Price');
            $spreadsheet->getActiveSheet()->setCellValue("B{$k}", $fabname[$j]);
            $spreadsheet->getActiveSheet()->setCellValue("C{$k}", $costall['costp2']['fab']['a'.$v]);

        }else{
            $spreadsheet->getActiveSheet()->setCellValue("B{$k}", $fabname[$j]);
            $spreadsheet->getActiveSheet()->setCellValue("C{$k}", $costall['costp2']['fab']['a'.$v]);
        }


    }
    $spreadsheet->getActiveSheet()->getStyle("A{$k}")->getAlignment()->setWrapText(true);

    $k++;
    $v++;
}



unset($_SESSION['costall'] ); //注销SESSION

// Set active sheet index to the first sheet, so Excel opens this as the first sheet
    $spreadsheet->setActiveSheetIndex(0); //返回第一页

    $spreadsheet->getActiveSheet()->getPageMargins()->setRight(0.1); //设置打印边距
    $spreadsheet->getActiveSheet()->getPageMargins()->setLeft(0.1); //

$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页
$output=  ($_GET['action'] == 'formprint' )? 1:0;
//$output= 0;
$nt = date("YmdHis",time()); //转换为日期。
$filenameout = 'costallout'.$nt.'.xlsx';
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
