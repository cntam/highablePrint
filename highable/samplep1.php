<?php
session_start();

require_once('autoloadconfig.php');  //判断是否在线

require_once ('img.php');

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器




$samplep1 =   $_SESSION['samplep1'];
//var_dump($samplep1);

//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/samplep1.xlsx');
$sheet = $spreadsheet->getActiveSheet();

$spreadsheet->getActiveSheet()->setTitle("第一页");

$spreadsheet->getDefaultStyle()->getFont()->setName('微软雅黑');
$spreadsheet->getDefaultStyle()->getFont()->setSize(12);
//$spreadsheet->getActiveSheet()->getDefaultRowDimension()->setRowHeight(50);

$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(15);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(10);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(15);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(15);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(15);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(15);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(10);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('H')->setWidth(15);  //列宽度


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
	
//    'borders' => [
//        'top' => [
//            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
//        ],
//		'bottom' => [
//            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
//        ],
//		'left' => [
//            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
//        ],
//		'right' => [
//            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
//        ],
//    ],
   
];


//$spreadsheet->getActiveSheet()->setCellValue('C4', $samplep1["alist"]['a1']);
//$spreadsheet->getActiveSheet()->setCellValue('H4', $samplep1["alist"]['a2']);

for($i=1,$u=0;$i <= 6 ;$i++,$u++){
    $arownum = 3 + $i;
    $a1 = 'a'. (1 + ($u * 2)) ;
    $a2 = 'a'. (2 + ($u * 2)) ;

    $spreadsheet->getActiveSheet()->setCellValue('C'.$arownum, $samplep1["alist"][$a1]);
    $spreadsheet->getActiveSheet()->setCellValue('H'.$arownum, $samplep1["alist"][$a2]);
}



/**
 * 图片模块
 */

$img = $samplep1["alist"]['a13'];
if ($img == '') {
    $haveimg = false;  //没有图片

} else {

    $path = $img;
    $pathinfo = pathinfo($path);
    //echo "扩展名：$pathinfo[extension]";

    if ($pathinfo['extension'] == 'pdf') {

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
    $drawing->setName('FABRIC RECODE');
    $drawing->setDescription('FABRIC RECODE');
//$drawing->setImageResource($gdImage);
    $drawing->setImageResource($img);
    $drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
    $drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
//$drawing->setHeight($width);

//$drawing->setHeight($width>550 ? 550:$width);
    $drawing->setHeight($height>300 ? 300:$height);
//$drawing->setHeight(150);


    $drawing->setCoordinates("A12");
    $drawing->setOffsetX(5);
    $drawing->setOffsetY(5);
    $drawing->setWorksheet($spreadsheet->getActiveSheet());
}
/* 图片模块 */


//特殊工序
if(is_array($samplep1['b10name']) && (count($samplep1['b10name']) >0 )){

    $b10 = '';
    foreach ($samplep1['b10name'] as $value){
        $b10 .= '  '.$value;
    }

}

$spreadsheet->getActiveSheet()->setCellValue('B33', $b10);
$spreadsheet->getActiveSheet()->setCellValue('B35', $samplep1['data']["blist"]['b11']);
$spreadsheet->getActiveSheet()->setCellValue('B37', $samplep1['data']["blist"]['b12']);


for($i=1,$u=0;$i <= $samplep1['data']["clist"]["clistnum"] ;$i++,$u++){
    $arownum = 40 + $u;

    $spreadsheet->getActiveSheet()->setCellValue('B'.$arownum, $samplep1['data']["clist"]['c1'][$u]);
    $spreadsheet->getActiveSheet()->setCellValue('C'.$arownum, $samplep1['data']["clist"]['c2'][$u]);
    $spreadsheet->getActiveSheet()->setCellValue('D'.$arownum, $samplep1['data']["clist"]['c3'][$u]);
    $spreadsheet->getActiveSheet()->setCellValue('E'.$arownum, $samplep1['data']["clist"]['c4'][$u]);
    $spreadsheet->getActiveSheet()->setCellValue('F'.$arownum, $samplep1['data']["clist"]['c5'][$u]);
    $spreadsheet->getActiveSheet()->setCellValue('G'.$arownum, $samplep1['data']["clist"]['c6'][$u]);
}
$spreadsheet->getActiveSheet()->setCellValue('H42', $samplep1['data']["clist"]['c7']);

for($i=1,$u=0;$i <= $samplep1['data']["dlist"]["dlistnum"] ;$i++,$u++){
    $arownum = 46 + $u;

    $spreadsheet->getActiveSheet()->setCellValue('A'.$arownum, $samplep1['data']["dlist"]['d1'][$u]);
    $spreadsheet->getActiveSheet()->setCellValue('B'.$arownum, $samplep1['data']["dlist"]['d2'][$u]);
    $spreadsheet->getActiveSheet()->setCellValue('C'.$arownum, $samplep1['data']["dlist"]['d3'][$u]);
    $spreadsheet->getActiveSheet()->setCellValue('D'.$arownum, $samplep1['data']["dlist"]['d4'][$u]);
    $spreadsheet->getActiveSheet()->setCellValue('E'.$arownum, $samplep1['data']["dlist"]['d5'][$u]);
    $spreadsheet->getActiveSheet()->setCellValue('F'.$arownum, $samplep1['data']["dlist"]['d6'][$u]);
    $spreadsheet->getActiveSheet()->setCellValue('G'.$arownum, $samplep1['data']["dlist"]['d7'][$u]);
    $spreadsheet->getActiveSheet()->setCellValue('H'.$arownum, $samplep1['data']["dlist"]['d8'][$u]);
}

//

/**
 * 款号备注
 */

$titlearr = array('面布 ：',' 裡布1：',' 裡布2：','裡布3：','撞布1：','撞布2：','撞布3');
$f = 0;
$browrow = 25;
for($r=0;$r<=6;$r++){

    $b1 = 'b'. (1 + $r );
    $b1 = $samplep1['data']["blist"][$b1];
    if('on' == $b1){
        $spreadsheet->getActiveSheet()->setCellValue('B'.$browrow, $titlearr[$r]);
        $spreadsheet->getActiveSheet()->setCellValue('C'.$browrow, $samplep1['data']["blist"]["b1v"][$r]);
        $browrow++;
        $f++;
    }
}

if($samplep1['data']["blist"]["formnumb"] > 0){
    for($r=0 , $k = 1;$k<= $samplep1['data']["blist"]["formnumb"];$r++,$k++){

            $spreadsheet->getActiveSheet()->setCellValue('B'.$browrow, $samplep1['data']["blist"]["b8"][$r]);
            $spreadsheet->getActiveSheet()->setCellValue('C'.$browrow, $samplep1['data']["blist"]["b9"][$r]);
        $spreadsheet->getActiveSheet()->getStyle("B{$browrow}")->applyFromArray($styleArray);
            $browrow++;
            $f++;

if($f>7){
    $spreadsheet->getActiveSheet()->insertNewRowBefore($browrow, 1);
    $spreadsheet->getActiveSheet()->mergeCells("C{$browrow}:H{$browrow}");
    $spreadsheet->getActiveSheet()->getStyle("B{$browrow}")->applyFromArray($styleArray);
}
    }
}

//
$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页


/**
 * 第二页
 */
$spreadsheet->setActiveSheetIndex(1);  //設置當前活動表
$sheet = $spreadsheet->getActiveSheet();

$spreadsheet->getDefaultStyle()->getFont()->setName('微软雅黑');
$spreadsheet->getDefaultStyle()->getFont()->setSize(12);
//$spreadsheet->getActiveSheet()->getDefaultRowDimension()->setRowHeight(50);

$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(10);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(5);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(10);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(10);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(10);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(10);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(5);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('H')->setWidth(15);  //列宽度

for($i=1,$u=0;$i <= 6 ;$i++,$u++){
    $arownum = 3 + $i;
    $a1 = 'a'. (1 + ($u * 2)) ;
    $a2 = 'a'. (2 + ($u * 2)) ;

    $spreadsheet->getActiveSheet()->setCellValue('C'.$arownum, $samplep1["alist"][$a1]);
    $spreadsheet->getActiveSheet()->setCellValue('H'.$arownum, $samplep1["alist"][$a2]);
}
$spreadsheet->getActiveSheet()->setCellValue('H9', '2');

if($samplep1["samplep2"]['blist']['b1']){

    /**
     * 图片模块
     */

    $img = $samplep1["samplep2"]['blist']['b2'];
    if ($img == '') {
        $haveimg = false;  //没有图片

    } else {

        $path = $img;
        $pathinfo = pathinfo($path);
        //echo "扩展名：$pathinfo[extension]";

        if ($pathinfo['extension'] == 'pdf') {

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
        $drawing->setName('FABRIC RECODE');
        $drawing->setDescription('FABRIC RECODE');
//$drawing->setImageResource($gdImage);
        $drawing->setImageResource($img);
        $drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
        $drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
//$drawing->setHeight($width);

//$drawing->setHeight($width>550 ? 550:$width);
        $drawing->setHeight($height>300 ? 300:$height);
//$drawing->setHeight(150);

        $drawing->setCoordinates("C11");
        $drawing->setOffsetX(5);
        $drawing->setOffsetY(5);
        $drawing->setWorksheet($spreadsheet->getActiveSheet());
    }
    /* 图片模块 */

}else{
    $spreadsheet->getActiveSheet()->mergeCells("C11:H20");
    $spreadsheet->getActiveSheet()->getStyle('C11:H20')->getAlignment()->setWrapText(true);  //在单元格中写入换行符“\ n”（ALT +“Enter”）
    $spreadsheet->getActiveSheet()->getStyle("C11:H20")->applyFromArray($styleArray);
    $spreadsheet->getActiveSheet()->setCellValue('C11', $samplep1["samplep2"]['comment']);
}
$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$spreadsheet->setActiveSheetIndex(0);

//unset($_SESSION['samplep1'] ); //注销SESSION

$output=  ($_GET['action'] == 'formdown' )? 1:0;
$nt = date("YmdHis",time()); //转换为日期。
$filenameout = 'samplep1out'.$nt.'.xlsx';
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
