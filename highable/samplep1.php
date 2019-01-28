<?php
require_once 'aidenfunc.php';

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器

$samplep1 =   $_SESSION['samplep1'];

//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/samplep1.xlsx');
$sheet = $spreadsheet->getActiveSheet();

$spreadsheet->getActiveSheet()->setTitle("sheet1");

$spreadsheet->getDefaultStyle()->getFont()->setName('微软雅黑');
$spreadsheet->getDefaultStyle()->getFont()->setSize(12);
//$spreadsheet->getActiveSheet()->getDefaultRowDimension()->setRowHeight(50);

$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(15);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(15);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(15);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(15);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(15);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(15);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(10);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('H')->setWidth(15);  //列宽度


$spreadsheet->getActiveSheet()->getPageMargins()->setRight(0.1);
$spreadsheet->getActiveSheet()->getPageMargins()->setLeft(0.1); //页边距


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


    $drawing->setWidthAndHeight(480,340);  //设置图片最大宽度 高度

//    $drawing->setWidth(100);
//$drawing->setHeight($width>550 ? 550:$width);
    //$drawing->setHeight($height>300 ? 300:$height);
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
$spreadsheet->getActiveSheet()->getStyle("B33:D33")->applyFromArray($noborderTopLeft);
$spreadsheet->getActiveSheet()->setCellValue('B35', $samplep1['data']["blist"]['b11']);
$spreadsheet->getActiveSheet()->getStyle("B35:D35")->applyFromArray($noborderTopLeft);
$spreadsheet->getActiveSheet()->setCellValue('B37', $samplep1['data']["blist"]['b12']);
$spreadsheet->getActiveSheet()->getStyle("B37:D37")->applyFromArray($noborderTopLeft);



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

/**
 * 色码分配
 */

$spreadsheet->getActiveSheet()->setCellValue('H43', $samplep1['data']["clist"]['clistresult']); //最后结果
for($i=1,$u=0;$u <= $samplep1['data']["clist"]["clistrow"] ;$i++,$u++){
    $arownum = 40 + $u;

    if($u > 2){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($arownum, 1);
    }

    $c1 = stripcslashes($samplep1['data']["clist"]['c1'][$u]);
    $spreadsheet->getActiveSheet()->setCellValue('A'.$arownum, $c1);
    $spreadsheet->getActiveSheet()->setCellValue('B'.$arownum, $samplep1['data']["clist"]['c2'][$u]);
    $spreadsheet->getActiveSheet()->setCellValue('C'.$arownum, $samplep1['data']["clist"]['c3'][$u]);
    $spreadsheet->getActiveSheet()->setCellValue('D'.$arownum, $samplep1['data']["clist"]['c4'][$u]);
    $spreadsheet->getActiveSheet()->setCellValue('E'.$arownum, $samplep1['data']["clist"]['c5'][$u]);
    $spreadsheet->getActiveSheet()->setCellValue('F'.$arownum, $samplep1['data']["clist"]['c6'][$u]);
    $spreadsheet->getActiveSheet()->setCellValue('G'.$arownum, $samplep1['data']["clist"]['c7'][$u]);
    $spreadsheet->getActiveSheet()->setCellValue('H'.$arownum, $samplep1['data']["clist"]['c8'][$u]);

}
$spreadsheet->getActiveSheet()->setCellValue('H40', '总计');
/**
 * 色码分配
 */

/**
 * 款号备注
 */

$titlearr = array('面布 ：','裡布1：','裡布2：','裡布3：','撞布1：','撞布2：','撞布3');
$f = 0;
$browrow = 25;
for($r=0;$r<=6;$r++){

    $b1 = 'b'. (1 + $r );
    $b1 = $samplep1['data']["blist"][$b1];
    if('on' == $b1){
        $spreadsheet->getActiveSheet()->setCellValue('B'.$browrow, $titlearr[$r]);
        $spreadsheet->getActiveSheet()->setCellValue('C'.$browrow, $samplep1['data']["blist"]["b1v"][$r]);
        $spreadsheet->getActiveSheet()->getStyle("C{$browrow}:H{$browrow}")->applyFromArray($noborderTopLeft);
        $browrow++;
        $f++;
    }
}

if($samplep1['data']["blist"]["formnumb"] > 0){
    for($r=0 , $k = 1;$k<= $samplep1['data']["blist"]["formnumb"];$r++,$k++){

            $spreadsheet->getActiveSheet()->setCellValue('B'.$browrow, $samplep1['data']["blist"]["b8"][$r]);
            $spreadsheet->getActiveSheet()->setCellValue('C'.$browrow, $samplep1['data']["blist"]["b9"][$r]);
        $spreadsheet->getActiveSheet()->getStyle("B{$browrow}")->applyFromArray($noborderTopLeft);
        $spreadsheet->getActiveSheet()->getStyle("C{$browrow}:H{$browrow}")->applyFromArray($noborderTopLeft);
            $browrow++;
            $f++;

if($f>7){
    $spreadsheet->getActiveSheet()->insertNewRowBefore($browrow, 1);
    $spreadsheet->getActiveSheet()->mergeCells("C{$browrow}:H{$browrow}");
    $spreadsheet->getActiveSheet()->getStyle("B{$browrow}")->applyFromArray($noborderTopLeft);
    $spreadsheet->getActiveSheet()->getStyle("C{$browrow}:H{$browrow}")->applyFromArray($noborderTopLeft);
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


$sheet->getColumnDimension('A')->setWidth(15);  //列宽度
$sheet->getColumnDimension('B')->setWidth(15);  //列宽度
$sheet->getColumnDimension('C')->setWidth(15);  //列宽度
$sheet->getColumnDimension('D')->setWidth(15);  //列宽度
$sheet->getColumnDimension('E')->setWidth(15);  //列宽度
$sheet->getColumnDimension('F')->setWidth(15);  //列宽度
$sheet->getColumnDimension('G')->setWidth(10);  //列宽度
$sheet->getColumnDimension('H')->setWidth(15);  //列宽度


$sheet->getPageMargins()->setRight(0.1);
$sheet->getPageMargins()->setLeft(0.1); //页边距

$spreadsheet->getActiveSheet()->setTitle("sheet2");
$spreadsheet->getDefaultStyle()->getFont()->setName('微软雅黑');
$spreadsheet->getDefaultStyle()->getFont()->setSize(12);
//$spreadsheet->getActiveSheet()->getDefaultRowDimension()->setRowHeight(50);


for($i=1,$u=0;$i <= 6 ;$i++,$u++){
    $arownum = 3 + $i;
    $a1 = 'a'. (1 + ($u * 2)) ;
    $a2 = 'a'. (2 + ($u * 2)) ;

    $spreadsheet->getActiveSheet()->setCellValue('C'.$arownum, $samplep1["alist"][$a1]);
    $spreadsheet->getActiveSheet()->setCellValue('H'.$arownum, $samplep1["alist"][$a2]);
}
$spreadsheet->getActiveSheet()->setCellValue('H9', '2');






    /**
     * p2 评语
     */
    if(empty($samplep1["samplep2"]['blist']['b2'])){
        $row = 11;
        $ishaveimg = false;
    }else{
        $row = 25;
        $ishaveimg = true;
    }

    if($ishaveimg){
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

            $drawing->setWidthAndHeight(520,360);  //设置图片最大宽度 高度
//$drawing->setHeight($width>550 ? 550:$width);
            //$drawing->setHeight($height>300 ? 300:$height);
//$drawing->setHeight(150);

            $drawing->setCoordinates("C11");
            $drawing->setOffsetX(5);
            $drawing->setOffsetY(5);
            $drawing->setWorksheet($spreadsheet->getActiveSheet());
        }
        /* 图片模块 */
    }




    foreach ($samplep1["samplep2"]['commentarr'] as $item => $value){
    $spreadsheet->getActiveSheet()->mergeCells("C{$row}:H{$row}");
    $spreadsheet->getActiveSheet()->getStyle("C{$row}:H{$row}")->getAlignment()->setWrapText(true);  //在单元格中写入换行符“\ n”（ALT +“Enter”）
    $spreadsheet->getActiveSheet()->getStyle("C{$row}:H{$row}")->applyFromArray($noborderTopLeft);
    $spreadsheet->getActiveSheet()->setCellValue('C'.$row, $value);
    $row++;
    }


//    $spreadsheet->getActiveSheet()->mergeCells("B25:H48");
//    $spreadsheet->getActiveSheet()->getStyle('B25:H48')->getAlignment()->setWrapText(true);  //在单元格中写入换行符“\ n”（ALT +“Enter”）
//    $spreadsheet->getActiveSheet()->getStyle("B25:H48")->applyFromArray($noborderTopLeft);
//    $spreadsheet->getActiveSheet()->setCellValue('B25', $samplep1["samplep2"]['comment']);
    /**
     * p2 评语
     */
$sheet->getPageSetup()->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);//打印纸张 A4
$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$spreadsheet->setActiveSheetIndex(0);

unset($_SESSION['samplep1'] ); //注销SESSION

$filenameout = "SampleOrder_{$samplep1['clientname']}_{$samplep1['orderno']}_";
outExcel($spreadsheet,$filenameout);
