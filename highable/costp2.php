<?php
session_start();

$costp2 =  $_SESSION['costp2'];

require_once('autoloadconfig.php');  //判断是否在线

require_once ('img.php');

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

//var_dump($costp2);
$temno = $costp2["temno"];
$titlearr = unserialize(gzuncompress(base64_decode($costp2["cctitle"])));
//print_r($titlearr);

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

$spreadsheet->getActiveSheet()->setTitle("COST CHART sheet1");
//$sheet->setCellValue('A1', 'Hello World !');
$spreadsheet->getDefaultStyle()->getFont()->setName('微软雅黑');
$spreadsheet->getDefaultStyle()->getFont()->setSize(12);
//$spreadsheet->getActiveSheet()->getDefaultRowDimension()->setRowHeight(50); //行默认高度
$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(50);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(50);  //列宽度

$spreadsheet->getActiveSheet()->getRowDimension('1')->setRowHeight(36); //列高度
$spreadsheet->getActiveSheet()->getRowDimension('2')->setRowHeight(50); //列高度

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


//$daftitle=array("CLIENT:","Sketch","Style no.：");
//$daftitlenum = count($daftitle);
//for ($j=0,$k = 1,$v =1 ;$j<$daftitlenum;$j++,$k++){
//
//    $spreadsheet->getActiveSheet()->setCellValue("A{$k}", $daftitle[$j]);
//    $spreadsheet->getActiveSheet()->getStyle("A{$k}")->applyFromArray($styleArray);
//    $spreadsheet->getActiveSheet()->getStyle("A{$k}")->getAlignment()->setWrapText(true);
//
//
//
//}

$spreadsheet->getActiveSheet()->setCellValue("A1", 'CLIENT:');
$spreadsheet->getActiveSheet()->getStyle("A1")->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle("A1")->getAlignment()->setWrapText(true);

$spreadsheet->getActiveSheet()->setCellValue("A2", 'Sketch');
$spreadsheet->getActiveSheet()->getStyle("A2:A8")->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle("A2:A8")->getAlignment()->setWrapText(true);

$spreadsheet->getActiveSheet()->setCellValue("A9", 'Style no.：');
$spreadsheet->getActiveSheet()->getStyle("A9")->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle("A9")->getAlignment()->setWrapText(true);
$spreadsheet->getActiveSheet()->getStyle("B2:B8")->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle("B9")->applyFromArray($styleArray);


$spreadsheet->getActiveSheet()->setCellValue("B1", $costp2["clientname"]);
$spreadsheet->getActiveSheet()->getStyle("B1")->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle("B1")->getAlignment()->setWrapText(true);

$spreadsheet->getActiveSheet()->setCellValue("C1", $costp2["alist"]["a1"]);
$spreadsheet->getActiveSheet()->getStyle("C1")->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle("C1")->getAlignment()->setWrapText(true);


$spreadsheet->getActiveSheet()->setCellValue("C2", $costp2["alist"]["a3"]);
$spreadsheet->getActiveSheet()->getStyle("C2:C8")->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle("C2:C8")->getAlignment()->setWrapText(true);

$spreadsheet->getActiveSheet()->mergeCells("C4:C6");
$spreadsheet->getActiveSheet()->setCellValue("C4", $costp2["alist"]["a4"]);

$spreadsheet->getActiveSheet()->getStyle("C4")->getAlignment()->setWrapText(true);

$spreadsheet->getActiveSheet()->setCellValue("C8", $costp2["alist"]["a5"]);
$spreadsheet->getActiveSheet()->getStyle("C9")->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle("C8")->getAlignment()->setWrapText(true);

$spreadsheet->getActiveSheet()->setCellValue("B9", $costp2["styleno"]);
$spreadsheet->getActiveSheet()->getStyle("B9")->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle("B9")->getAlignment()->setWrapText(true);





/**
 * 图片模块
 */

$img = $costp2["alist"]["a2"];
if ($img == '') {
    $haveimg = false;  //没有图片

} else {

    $path = $img;
    $pathinfo = pathinfo($path);
    //echo "扩展名：$pathinfo[extension]";

    if ($pathinfo["extension"] == 'pdf') {

        $img = pdficon();
        $haveimg = true;
    } else {
        $haveimg = true;
    }
}


if ($haveimg){
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
    $width = imagesx($img);
    $height = imagesy($img);


// Add a drawing to the worksheet
    $drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
    $drawing->setName('img');
    $drawing->setDescription('img');
//$drawing->setImageResource($gdImage);
    $drawing->setImageResource($img);
    $drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
    $drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
//$drawing->setHeight($width);

    $drawing->setWidth($width>250 ? 250:$width);
    //$drawing->setHeight($height>130 ? 130:$height);
//$drawing->setHeight(150);


    //$drawing->setCoordinates($cola.'2');
    $drawing->setCoordinates('B2');
    $drawing->setOffsetX(5);
    $drawing->setOffsetY(5);
    $drawing->setWorksheet($spreadsheet->getActiveSheet());
}
/* 图片模块 */



switch ($temno){
     case 7:  //MCQ
     case 8:  //PS

     $tdefault = true;
      break;

    default:

    break;
    }

    for($i=1,$v = 0,$l = 10;$i<= count($titlearr);$i++,$v++,$l++){

        $spreadsheet->getActiveSheet()->setCellValue("A{$l}", $titlearr[$v]);
        $spreadsheet->getActiveSheet()->getStyle("A{$l}")->applyFromArray($styleArray);
        $spreadsheet->getActiveSheet()->getStyle("A{$l}")->getAlignment()->setWrapText(true);

        $spreadsheet->getActiveSheet()->setCellValue("B{$l}", $costp2["clist"]["c".$i]);
        $spreadsheet->getActiveSheet()->getStyle("B{$l}")->applyFromArray($styleArray);
        $spreadsheet->getActiveSheet()->getStyle("B{$l}")->getAlignment()->setWrapText(true);

        $spreadsheet->getActiveSheet()->setCellValue("C{$l}", $costp2["dlist"]["d".$i]);
        $spreadsheet->getActiveSheet()->getStyle("C{$l}")->applyFromArray($styleArray);
        $spreadsheet->getActiveSheet()->getStyle("C{$l}")->getAlignment()->setWrapText(true);
        }

/**
 * 中间布料栏
 */

if($costp2["alist"]["a10"] > 0){    //如果行数大于12 增加行
    $spreadsheet->getActiveSheet()->setCellValue("D1", $costp2["alist"]["a10"]);
    $addlist = 10;

    for($n = 1,$r=($costp2["alist"]["a10"]-1);$n<=($costp2["alist"]["a10"]);$n++,$r-- ){

        $spreadsheet->getActiveSheet()->insertNewRowBefore($addlist, 3);

        /**第一行*/
        $spreadsheet->getActiveSheet()->setCellValue("A10", $costp2["alist"]["a11"][$r]);
        $spreadsheet->getActiveSheet()->getStyle("A10")->applyFromArray($styleArray);
        $spreadsheet->getActiveSheet()->getStyle("A10")->getAlignment()->setWrapText(true);

        $spreadsheet->getActiveSheet()->setCellValue("B10", $costp2["alist"]["a12"][$r]);
        $spreadsheet->getActiveSheet()->getStyle("B10")->applyFromArray($styleArray);
        $spreadsheet->getActiveSheet()->getStyle("B10")->getAlignment()->setWrapText(true);

        $spreadsheet->getActiveSheet()->setCellValue("C10", $costp2["alist"]["a13"][$r]);
        $spreadsheet->getActiveSheet()->getStyle("C10")->applyFromArray($styleArray);
        $spreadsheet->getActiveSheet()->getStyle("C10")->getAlignment()->setWrapText(true);
        /**第一行*/

        /**第2行*/
        $spreadsheet->getActiveSheet()->setCellValue("A11", $costp2["alist"]["a14"][$r]);
        $spreadsheet->getActiveSheet()->getStyle("A11")->applyFromArray($styleArray);
        $spreadsheet->getActiveSheet()->getStyle("A11")->getAlignment()->setWrapText(true);

        if(1 == $costp2["alist"]["a15"][$r]){
            $alista15 = 'USD$';
        }elseif (2 == $costp2["alist"]["a15"][$r]){
            $alista15 = 'HKD$';
        }elseif (3 == $costp2["alist"]["a15"][$r]){
            $alista15 = 'RMB￥';
        }

        $alistB11 = $alista15 .' '. $costp2["alist"]["a16"][$r] .' '. $costp2["alist"]["a17"][$r] ;

        $spreadsheet->getActiveSheet()->setCellValue("B11", $alistB11);
        $spreadsheet->getActiveSheet()->getStyle("B11")->applyFromArray($styleArray);
        $spreadsheet->getActiveSheet()->getStyle("B11")->getAlignment()->setWrapText(true);

        $spreadsheet->getActiveSheet()->setCellValue("C11", $costp2["alist"]["a18"][$r]);
        $spreadsheet->getActiveSheet()->getStyle("C11")->applyFromArray($styleArray);
        $spreadsheet->getActiveSheet()->getStyle("C11")->getAlignment()->setWrapText(true);
        /**第2行*/

        /**第3行*/
        $spreadsheet->getActiveSheet()->setCellValue("A12", $costp2["alist"]["a19"][$r]);
        $spreadsheet->getActiveSheet()->getStyle("A12")->applyFromArray($styleArray);
        $spreadsheet->getActiveSheet()->getStyle("A12")->getAlignment()->setWrapText(true);

        if(1 == $costp2["alist"]["a22"][$r]){
            $alista22 = 'y/DZ';
        }elseif (2 == $costp2["alist"]["a22"][$r]){
            $alista22 = 'y/PC';
        }

        $alistB12 =   $costp2["alist"]["a20"][$r] .' '. $costp2["alist"]["a21"][$r] .' '. $alista22;

        $spreadsheet->getActiveSheet()->setCellValue("B12", $alistB12);
        $spreadsheet->getActiveSheet()->getStyle("B12")->applyFromArray($styleArray);
        $spreadsheet->getActiveSheet()->getStyle("B12")->getAlignment()->setWrapText(true);

        if(1 == $costp2["alist"]["a23"][$r]){
            $alista23 = 'y/DZ';
        }elseif (2 == $costp2["alist"]["a23"][$r]){
            $alista23 = 'y/PC';
        }

        $alistB12 =   $costp2["alist"]["a20"][$r] .' '. $costp2["alist"]["a21"][$r] .' '. $alista22;

        $spreadsheet->getActiveSheet()->setCellValue("C12", $costp2["alist"]["a23"][$r]);
        $spreadsheet->getActiveSheet()->getStyle("C12")->applyFromArray($styleArray);
        $spreadsheet->getActiveSheet()->getStyle("C12")->getAlignment()->setWrapText(true);
        /**第3行*/
    }
}


/**
 * //中间布料栏
 */


/**
 *  下面就是 旧的
 */


//for ($i = 1;$i<44;$i++) {
//    $spreadsheet->getActiveSheet()->getStyle("A{$i}")->applyFromArray($styleArray);
//}
//
//
//$styleArray2 = [
//
//    'alignment' => [
//        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
//        'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
//    ],
//
//    'borders' => [
//        'top' => [
//            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
//        ],
//        'bottom' => [
//            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
//        ],
//        'left' => [
//            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
//        ],
//        'right' => [
//            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
//        ],
//    ],
//
//];
//for ($i = 1;$i<44;$i++) {
//
//    $spreadsheet->getActiveSheet()->getStyle("B{$i}")->applyFromArray($styleArray2);
//    $spreadsheet->getActiveSheet()->getStyle("C{$i}")->applyFromArray($styleArray2);
//}
///*
//$spreadsheet->getActiveSheet()->getStyle('C2:D2')->applyFromArray($styleArray);
//$spreadsheet->getActiveSheet()->getStyle('A3:B3')->applyFromArray($styleArray);
//$spreadsheet->getActiveSheet()->getStyle('C3:D3')->applyFromArray($styleArray);
//$spreadsheet->getActiveSheet()->getStyle('A4:B4')->applyFromArray($styleArray);
//$spreadsheet->getActiveSheet()->getStyle('C4:D4')->applyFromArray($styleArray);
//$spreadsheet->getActiveSheet()->getStyle('A5:B5')->applyFromArray($styleArray);
//$spreadsheet->getActiveSheet()->getStyle('C5:D5')->applyFromArray($styleArray);
//*/
//$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(32); //列宽度
//$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(21);
//$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(32);
///*
//
//
//$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(19);
//
//$spreadsheet->getActiveSheet()->getRowDimension('1')->setRowHeight(40);
//$spreadsheet->getActiveSheet()->getRowDimension('2')->setRowHeight(160);
//$spreadsheet->getActiveSheet()->getRowDimension('3')->setRowHeight(160);
//$spreadsheet->getActiveSheet()->getRowDimension('4')->setRowHeight(160);
//$spreadsheet->getActiveSheet()->getRowDimension('5')->setRowHeight(160);
//*/
//
//// Set cell A1 with a string value
//
////$spreadsheet->getActiveSheet()->getStyle('A1:D1')->getFont()->setSize(14);
//
//
//$spreadsheet->getActiveSheet()->setCellValue('B1', $costp2['costname']);
//$spreadsheet->getActiveSheet()->setCellValue('A2', "DATE:");
//$spreadsheet->getActiveSheet()->setCellValue('B2', $costp2['costdata']);
//$spreadsheet->getActiveSheet()->setCellValue('A3', "款式：");
//$spreadsheet->getActiveSheet()->setCellValue('B3', $costp2['costno']);
//
//
//
//$img = $costp2['remarkimg2'];
//preg_match ('/.(jpg|gif|bmp|jpeg|png)/i', $img, $imgformat);
//$imgformat = $imgformat[1];
//switch ($imgformat)
//{
//    case "jpg":
//    case "jpeg":
//        $img = imagecreatefromjpeg($img);
//        break;
//    case "bmp":
//        $img =  imagecreatefromwbmp($img);
//        break;
//    case "gif":
//        $img =  imagecreatefromgif($img);
//        break;
//    case "png":
//        $img =   imagecreatefrompng($img);
//        break;
//}
//
////$img = imagecreatefromjpeg($img);
//
//$width = imagesx($img);
//
//$height = imagesy($img);
//
//
//// Generate an image
////$gdImage = @imagecreatetruecolor($width, $height) or die('Cannot Initialize new GD image stream');
////$textColor = imagecolorallocate($gdImage, 255, 255, 255);
////imagestring($gdImage, 1, 5, 5,  'Created with PhpSpreadsheet', $textColor);
//
//// Add a drawing to the worksheet
//$drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
//$drawing->setName($costp2['costname']);
//$drawing->setDescription($costp2['costname']);
////$drawing->setImageResource($gdImage);
//$drawing->setImageResource($img);
//$drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
//$drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
////$drawing->setHeight($width);
//
////$drawing->setHeight($width>550 ? 550:$width);
////$drawing->setWidth(300);
//$drawing->setHeight(135);
//$drawing->setCoordinates('B4');
//$drawing->setOffsetX(130);
//$drawing->setOffsetY(5);
//$drawing->setWorksheet($spreadsheet->getActiveSheet());
//
//
//
//$spreadsheet->getActiveSheet()->setCellValue('A12', "Style no：");
//$spreadsheet->getActiveSheet()->setCellValue('B12', $costp2['styleno']);
//
//
//$fabname=array("SLEF FABRIC：","Fabric Cost：","Cons./Doz(NET)：",'CONTRAST 1 fabric：','Fabric Cost：','Cons./Doz(NET)：','CONTRAST 2 fabric：','Fabric Cost：','Cons./Doz(NET)：','CONTRAST 3 fabric：','Fabric Cost：','Cons./Doz(NET)：','FABRIC COST：','Interlining @ 15：','Thread：','MCQ label(main label,size label & CO label)：','Carton：','MCQ poly bag & hangtag：','Fabric test cost：','Sticker：','18L Shell button(7+1) use for centre front placker：','16L Shell button(2+1) use for cuff：','trimming cost：','Tatal trim cost(10%)：','Sewing(RMB:120.0/PC)：','Cut,Trim,Pack etc.','Factory Overhead','Profit margin','100-200PCS+CM30%','201-400PCS+CM15%','OVER 400PCS');
//$fabcou = count($fabname);
//for ($j=0,$k = 13,$v =1 ;$j<$fabcou;$j++){
//
//    $m = $k+2;
//    if($v<29){
//        $spreadsheet->getActiveSheet()->setCellValue("A{$k}", $fabname[$j]);
//        $spreadsheet->getActiveSheet()->setCellValue("B{$k}", $costp2['fab']['a'.$v]);
//    }else{
//        if($v == 29){
//            $spreadsheet->getActiveSheet()->mergeCells("A{$k}:A{$m}");
//            $spreadsheet->getActiveSheet()->setCellValue("A{$k}", 'Unit Price');
//            $spreadsheet->getActiveSheet()->setCellValue("B{$k}", $fabname[$j]);
//            $spreadsheet->getActiveSheet()->setCellValue("C{$k}", $costp2['fab']['a'.$v]);
//
//        }else{
//            $spreadsheet->getActiveSheet()->setCellValue("B{$k}", $fabname[$j]);
//            $spreadsheet->getActiveSheet()->setCellValue("C{$k}", $costp2['fab']['a'.$v]);
//        }
//
//
//    }
//    $spreadsheet->getActiveSheet()->getStyle("A{$k}")->getAlignment()->setWrapText(true);
//
//    $k++;
//    $v++;
//}

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

//unset($_SESSION['costp2'] ); //注销SESSION

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