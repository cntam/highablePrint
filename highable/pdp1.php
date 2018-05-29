<?php
session_start();
header("Content-type: text/html; charset=utf-8");
//require '../vendor/autoload.php';
//require '/home/pan/vendor/autoload.php';
require_once('autoloadconfig.php');  //判断是否在线

if($online){
    require_once '/home/pan/vendor/autoload.php';

}else{
    require_once '/Applications/XAMPP/xamppfiles/htdocs/composer/vendor/autoload.php';
}

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;

$pdp1 =  $_SESSION['pdp1'];
//var_dump($pdp1);

//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/rdp1.xlsx');
\PhpOffice\PhpSpreadsheet\Cell\Cell::setValueBinder( new \PhpOffice\PhpSpreadsheet\Cell\AdvancedValueBinder() );

$sheet = $spreadsheet->getActiveSheet();
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
$spreadsheet->getDefaultStyle()->getFont()->setSize(10);

$sheet->setCellValue('A3',  $pdp1["SPL_1_code"]);
$sheet->setCellValue('B3',  $pdp1["SPL_1_name"]);
$sheet->setCellValue('C3',  $pdp1["SPL_1_country"]);
$sheet->setCellValue('D3',  $pdp1["SPL_1_contact"]);
$sheet->setCellValue('E3',  $pdp1["SPL_1_address"]);
$sheet->setCellValue('F3',  'EMAIL:'.$pdp1["SPL_1_email"].' \n TEL:'.$pdp1["SPL_1_tel"].' \n MOBILE'.$pdp1["SPL_1_mobile"].'\n QQ:'.$pdp1["SPL_1_qq"]);
$sheet->setCellValue('G3',  $pdp1["SPL_1_goods"]);
$spreadsheet->getActiveSheet()->getStyle("F3")->getAlignment()->setWrapText(true);

$sheet->setCellValue('A4',  $pdp1["SPL_2_code"]);
$sheet->setCellValue('B4',  $pdp1["SPL_2_name"]);
$sheet->setCellValue('C4',  $pdp1["SPL_2_country"]);
$sheet->setCellValue('D4',  $pdp1["SPL_2_contact"]);
$sheet->setCellValue('E4',  $pdp1["SPL_2_address"]);
if($pdp1["SPL_2_code"]) {
    $sheet->setCellValue('F4', 'EMAIL:' . $pdp1["SPL_2_email"] . '\n TEL:' . $pdp1["SPL_2_tel"] . '\n MOBILE:' . $pdp1["SPL_2_mobile"] . ' \n QQ:' . $pdp1["SPL_2_qq"]);
}
$sheet->setCellValue('G4',  $pdp1["SPL_2_goods"]);
$spreadsheet->getActiveSheet()->getStyle("F4")->getAlignment()->setWrapText(true);

for($i = 5,$a = 0; $i<8  ;$i++){
    $col = chr(97 + $a);
    if($pdp1['spli35'][$col.'0']){
        $sheet->setCellValue("A{$i}", $pdp1['spli35'][$col.'0']);

        $sheet->setCellValue("B{$i}", $pdp1['spli35'][$col.'1']);

        $sheet->setCellValue("C{$i}", $pdp1['spli35'][$col.'2']);
        $sheet->setCellValue("D{$i}", $pdp1['spli35'][$col.'3']);
        $sheet->setCellValue("E{$i}", $pdp1['spli35'][$col.'4']);
        $sheet->setCellValue("F{$i}", 'EMAIL:'.$pdp1['spli35'][$col.'5'].' \n TEL:'.$pdp1['spli35'][$col.'6'].' \n MOBILE:'.$pdp1['spli35'][$col.'7'].' \n QQ:'.$pdp1['spli35'][$col.'8']);
        $sheet->setCellValue("G{$i}", $pdp1['spli35'][$col.'9']);
        $spreadsheet->getActiveSheet()->getStyle("F{$i}")->getAlignment()->setWrapText(true);
    }

    $a++;

}




            if ($pdp1['maxnum'] >= 0) {

                for ($formnum = 0; $formnum < $pdp1['maxnum']; $formnum++) {




                    /**
                     * FR 边框线
                     */
                    $styleArray = [
                        'borders' => [
                            'outline' => [
                                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
                                'color' => ['argb' => '00000000'],
                            ],
                        ],
                    ];

                    $anum = (10 + 18* $formnum);
                    $gnum = (26 + 18* $formnum);
                    $spreadsheet->getActiveSheet()->getStyle("A{$anum}:G{$gnum}")->applyFromArray($styleArray);
                    /* FR 边框线  */


                    /**
                     * FR 内容
                     */
                    $efnum = (10 + 18* $formnum);
                    $spreadsheet->getActiveSheet()->setCellValue("E{$efnum}", "DATE:");
                    $spreadsheet->getActiveSheet()->mergeCells("F{$efnum}:G{$efnum}");
                    $spreadsheet->getActiveSheet()->setCellValue("F{$efnum}", $pdp1["frlistcon"][$formnum][1]);


                    $efnum = (12 + 18* $formnum);
                    $spreadsheet->getActiveSheet()->setCellValue("E{$efnum}", "IHK NO.:");
                    $spreadsheet->getActiveSheet()->mergeCells("F{$efnum}:G{$efnum}");
                    $spreadsheet->getActiveSheet()->setCellValue("F{$efnum}", $pdp1["frlistcon"][$formnum][2]);

                    $efnum = (14 + 18* $formnum);
                    $spreadsheet->getActiveSheet()->setCellValue("E{$efnum}", "SUPPLIER:");
                    $spreadsheet->getActiveSheet()->mergeCells("F{$efnum}:G{$efnum}");
                    $spreadsheet->getActiveSheet()->setCellValue("F{$efnum}", $pdp1["frlistcon"][$formnum][3]);

                    $efnum = (16 + 18* $formnum);
                    $spreadsheet->getActiveSheet()->setCellValue("E{$efnum}", "SUPPLIER CODE:");
                    $spreadsheet->getActiveSheet()->mergeCells("F{$efnum}:G{$efnum}");
                    $spreadsheet->getActiveSheet()->setCellValue("F{$efnum}", $pdp1["frlistcon"][$formnum][4]);

                    $efnum = (18 + 18* $formnum);
                    $spreadsheet->getActiveSheet()->setCellValue("E{$efnum}", "COMP.:");
                    $spreadsheet->getActiveSheet()->mergeCells("F{$efnum}:G{$efnum}");
                    $spreadsheet->getActiveSheet()->setCellValue("F{$efnum}", $pdp1["frlistcon"][$formnum][5]);

                    $efnum = (20 + 18* $formnum);
                    $spreadsheet->getActiveSheet()->setCellValue("E{$efnum}", "WIDTH:");
                    $spreadsheet->getActiveSheet()->mergeCells("F{$efnum}:G{$efnum}");
                    $spreadsheet->getActiveSheet()->setCellValue("F{$efnum}", $pdp1["frlistcon"][$formnum][6]);

                    $efnum = (22 + 18* $formnum);
                    $spreadsheet->getActiveSheet()->setCellValue("E{$efnum}", "WEIGHT:");
                    $spreadsheet->getActiveSheet()->mergeCells("F{$efnum}:G{$efnum}");
                    $spreadsheet->getActiveSheet()->setCellValue("F{$efnum}", $pdp1["frlistcon"][$formnum][7]);

                    $efnum = (24 + 18* $formnum);
                    $efnum2 = (25 + 18* $formnum);
                    $spreadsheet->getActiveSheet()->setCellValue("E{$efnum}", "REMARK:");
                    $spreadsheet->getActiveSheet()->mergeCells("F{$efnum}:G{$efnum2}");
                    $spreadsheet->getActiveSheet()->setCellValue("F{$efnum}", $pdp1["frlistcon"][$formnum][8]);

                    $efnum = (26 + 18* $formnum);

                    $spreadsheet->getActiveSheet()->setCellValue("E{$efnum}", "PRICE:");
                    $spreadsheet->getActiveSheet()->mergeCells("F{$efnum}:G{$efnum}");
                    $spreadsheet->getActiveSheet()->setCellValue("F{$efnum}", $pdp1["frlistcon"][$formnum][10]);

                    /* FR 内容  */


                    /**
                     * 图片模块
                     */
                    $img = $pdp1["frlistcon"][$formnum][9];
                    if ($img == '') {
                        $haveimg = false;  //没有图片

                    } else {

                            $path = $img;
                            $pathinfo = pathinfo($path);
                            //echo "扩展名：$pathinfo[extension]";

                            if ($pathinfo['extension'] == 'pdf') {

                                $pdficon = 'data:image/png;base64,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';

                                $img = $pdficon;
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
    $drawing->setWidth(250);
//$drawing->setHeight(150);

    $afnum = (10 + 18* $formnum);
    $drawing->setCoordinates("A{$afnum}");
    $drawing->setOffsetX(5);
    $drawing->setOffsetY(5);
    $drawing->setWorksheet($spreadsheet->getActiveSheet());
}
                    /* 图片模块 */


                }   //for
};


//$img = $pdp1["FR_img"];
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
//$width = imagesx($img);
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
//$drawing->setName('FABRIC RECODE');
//$drawing->setDescription('FABRIC RECODE');
////$drawing->setImageResource($gdImage);
//$drawing->setImageResource($img);
//$drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
//$drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
////$drawing->setHeight($width);
//
////$drawing->setHeight($width>550 ? 550:$width);
//$drawing->setWidth(250);
////$drawing->setHeight(150);
//$drawing->setCoordinates('A10');
//$drawing->setOffsetX(5);
//$drawing->setOffsetY(5);
//$drawing->setWorksheet($spreadsheet->getActiveSheet());
//
///*$spreadsheet->getActiveSheet()
//    ->getColumnDimension('A')
//    ->setWidth(48);
//$spreadsheet->getActiveSheet()
//    ->getRowDimension(1)
//    ->setRowHeight(-1);*/
///*
//$spreadsheet->getActiveSheet()->getStyle("A".$listrow)
//    ->getAlignment()
//    ->setWrapText(true);
//$spreadsheet->getActiveSheet()->getStyle("A".$listrow)
//    ->getAlignment()
//    ->setShrinkToFit(true);
//*/
//$styleArray1 = [
//    'alignment' => [
//        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
//        'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP,
//        'wrapText' => true,
//        'ShrinkToFit'=>true,
//    ],
//    'font' => [
//        'Size' => '10',
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
//
//    ],
//
//];
//$spreadsheet->getActiveSheet()->getStyle("A3:G3")->applyFromArray($styleArray1);
//$spreadsheet->getActiveSheet()->getStyle("A4:G4")->applyFromArray($styleArray1);

//$spreadsheet->getActiveSheet()->getStyle("A".$listrow)->getFont()->setSize(8);


/**
 * SO 模块
 */
$SOnum = ($pdp1['maxnum'] > 0 ? ($pdp1['maxnum']-1) : 0);
$titlenum = (28 + 18 * ($pdp1['maxnum'] > 0 ? ($pdp1['maxnum']-1) : 0));

$spreadsheet->getActiveSheet()->setCellValue("A{$titlenum}", 'SAMPLE ORDER');

/**
 * FR 边框线
 */
$styleArrayso = [
    'borders' => [
        'outline' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
            'color' => ['argb' => '00000000'],
        ],
    ],
];

$anum = (29 + 18 * $SOnum );
$gnum = (50 + 18 * $SOnum );
$spreadsheet->getActiveSheet()->getStyle("A{$anum}:G{$gnum}")->applyFromArray($styleArrayso);
/* FR 边框线  */


$fgnum = (29 + 18 * $SOnum );
$spreadsheet->getActiveSheet()->setCellValue("E{$fgnum}", "DATE:");
$spreadsheet->getActiveSheet()->mergeCells("F{$fgnum}:G{$fgnum}");
$spreadsheet->getActiveSheet()->setCellValue("F{$fgnum}", $pdp1["SO_date"]);

$fgnum = (31 + 18 * $SOnum );
$spreadsheet->getActiveSheet()->setCellValue("E{$fgnum}", "CATEGORY:");
$spreadsheet->getActiveSheet()->mergeCells("F{$fgnum}:G{$fgnum}");
$spreadsheet->getActiveSheet()->setCellValue("F{$fgnum}", $pdp1["SO_category"]);

$fgnum = (33 + 18 * $SOnum );
$spreadsheet->getActiveSheet()->setCellValue("E{$fgnum}", "STYLE NO.:");
$spreadsheet->getActiveSheet()->mergeCells("F{$fgnum}:G{$fgnum}");
$spreadsheet->getActiveSheet()->setCellValue("F{$fgnum}", $pdp1["SO_styleno"]);

$fgnum = (35 + 18 * $SOnum );
$spreadsheet->getActiveSheet()->setCellValue("E{$fgnum}", "CLIENT:");
$spreadsheet->getActiveSheet()->mergeCells("F{$fgnum}:G{$fgnum}");
$spreadsheet->getActiveSheet()->setCellValue("F{$fgnum}", $pdp1["SO_client"]);

$fgnum = (37 + 18 * $SOnum );
$spreadsheet->getActiveSheet()->setCellValue("E{$fgnum}", "FABRIC:");
$spreadsheet->getActiveSheet()->mergeCells("F{$fgnum}:G{$fgnum}");
$spreadsheet->getActiveSheet()->setCellValue("F{$fgnum}", $pdp1["SO_fabric"]);

$fgnum = (39 + 18 * $SOnum );
$spreadsheet->getActiveSheet()->setCellValue("E{$fgnum}", "FABRIC INFO:");
$spreadsheet->getActiveSheet()->mergeCells("F{$fgnum}:G{$fgnum}");
$spreadsheet->getActiveSheet()->setCellValue("F{$fgnum}", $pdp1["SO_fabricinfo"]);

$fgnum = (41 + 18 * $SOnum );
$spreadsheet->getActiveSheet()->setCellValue("E{$fgnum}", "LINING:");
$spreadsheet->getActiveSheet()->mergeCells("F{$fgnum}:G{$fgnum}");
$spreadsheet->getActiveSheet()->setCellValue("F{$fgnum}", $pdp1["SO_lining"]);

$fgnum = (43 + 18 * $SOnum );
$spreadsheet->getActiveSheet()->setCellValue("E{$fgnum}", "LINING INFO:");
$spreadsheet->getActiveSheet()->mergeCells("F{$fgnum}:G{$fgnum}");
$spreadsheet->getActiveSheet()->setCellValue("F{$fgnum}", $pdp1["SO_lininginfo"]);

$fgnum = (45 + 18 * $SOnum );
$spreadsheet->getActiveSheet()->setCellValue("E{$fgnum}", "TRIM:");
$spreadsheet->getActiveSheet()->mergeCells("F{$fgnum}:G{$fgnum}");
$spreadsheet->getActiveSheet()->setCellValue("F{$fgnum}", $pdp1["SO_trim"]);

$fgnum = (47 + 18 * $SOnum );
$spreadsheet->getActiveSheet()->setCellValue("E{$fgnum}", "TRIM INFO:");
$spreadsheet->getActiveSheet()->mergeCells("F{$fgnum}:G{$fgnum}");
$spreadsheet->getActiveSheet()->setCellValue("F{$fgnum}", $pdp1["SO_triminfo"]);

$fgnum = (49 + 18 * $SOnum );
$fgnum2 = (50 + 18 * $SOnum );
$spreadsheet->getActiveSheet()->setCellValue("E{$fgnum}", "REMARK:");
$spreadsheet->getActiveSheet()->mergeCells("F{$fgnum}:G{$fgnum2}");
$spreadsheet->getActiveSheet()->setCellValue("F{$fgnum}", $pdp1["SO_remark"]);

/**
 * 图片模块
 */
$img = $pdp1["SO_img"];
if ($img == '') {
    $haveimg = false;  //没有图片

} else {

    $path = $img;
    $pathinfo = pathinfo($path);
    //echo "扩展名：$pathinfo[extension]";

    if ($pathinfo['extension'] == 'pdf') {

        $pdficon = 'data:image/png;base64,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';

        $img = $pdficon;
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
    $drawing->setWidth(280);
//$drawing->setHeight(150);

    $afnum = (29 + 18 * $SOnum);
    $drawing->setCoordinates("A{$afnum}");
    $drawing->setOffsetX(5);
    $drawing->setOffsetY(5);
    $drawing->setWorksheet($spreadsheet->getActiveSheet());
}
/* 图片模块 */

/* SO 模块 */
/*
$sheet->setCellValue("L18", $pdp1['fab4']); //裁法
$sheet->setCellValue("L22", $pdp1['fab4']); //针距如下
$sheet->setCellValue("L25", $pdp1['fab3']); //工艺说明及注意事项*/



//$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

unset($_SESSION['pdp1'] ); //注销SESSION

$output=  ($_GET['action'] == 'formdown' )? 1:0;
$nt = date("YmdHis",time()); //转换为日期。
$filenameout = 'rdp1out'.$nt.'.xlsx';
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
};

