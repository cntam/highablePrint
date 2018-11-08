<?php
session_start();

$cpsform =  $_SESSION['cpsform'];

require_once('autoloadconfig.php');  //判断是否在线

require_once ('img.php');

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

//var_dump($cpsform);
//$temno = $cpsform["temno"];
//$titlearr = unserialize(gzuncompress(base64_decode($cpsform["cctitle"])));
//print_r($titlearr);

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

$spreadsheet->getActiveSheet()->setTitle("CPS");
//$sheet->setCellValue('A1', 'Hello World !');
$spreadsheet->getDefaultStyle()->getFont()->setName('微软雅黑');
$spreadsheet->getDefaultStyle()->getFont()->setSize(12);
//$spreadsheet->getActiveSheet()->getDefaultRowDimension()->setRowHeight(50); //行默认高度
$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(20);  //列宽度

for($col=0;$col< count($cpsform['id']);$col++) {
    $Brow = chr(66 + $col * 2);  //B
    $Crow = chr(67 + $col * 2);  //C
    $spreadsheet->getActiveSheet()->getColumnDimension($Brow)->setWidth(30);  //列宽度
    $spreadsheet->getActiveSheet()->getColumnDimension($Crow)->setWidth(30);  //列宽度
}

$spreadsheet->getActiveSheet()->getRowDimension('1')->setRowHeight(36); //列高度
$spreadsheet->getActiveSheet()->getRowDimension('2')->setRowHeight(36); //列高度
$spreadsheet->getActiveSheet()->getRowDimension('3')->setRowHeight(160); //列高度

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

function getforexcate($forex) {
    switch ($forex){
        case 1:
            $output = 'USD$';
            break;
        case 2:
            $output = 'HKD$';
            break;
        case 3:
            $output = 'RMB￥';
            break;
        case 4:
            $output = 'EUR€';
            break;
        case 5:
            $output = 'JPY￥';
            break;

            default:
                 $output = 'USD$';
            break;
    }
    return $output;
    }

function isselect($value){
    if ( $value == 'on') {
        $output = '■  ';
    } else {
        $output = '□  ';
    }
    return $output;
}




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
/**
 * 头部
 */
for($col=0;$col< count($cpsform['id']);$col++){
    $Brow = chr(66 + $col * 2);  //B
    $Crow = chr(67 + $col * 2);  //C

    $spreadsheet->getActiveSheet()->setCellValue("A2", 'Sample order no.：');
    $spreadsheet->getActiveSheet()->getStyle("A2")->applyFromArray($styleArray);

    $spreadsheet->getActiveSheet()->setCellValue("A3", 'Sketch：');
    $spreadsheet->getActiveSheet()->getStyle("A3")->applyFromArray($styleArray);



    $BC = "{$Brow}1:{$Crow}1";
    $spreadsheet->getActiveSheet()->mergeCells($BC);
    $spreadsheet->getActiveSheet()->setCellValue($Brow.'1' , $cpsform['sampleorderno'][$col]);
    $spreadsheet->getActiveSheet()->getStyle($BC)->applyFromArray($styleArray);
    $spreadsheet->getActiveSheet()->getStyle($BC)->getAlignment()->setWrapText(true);

    $BC = "{$Brow}2:{$Crow}2";
    $spreadsheet->getActiveSheet()->mergeCells($BC);
    $spreadsheet->getActiveSheet()->setCellValue($Brow.'2' , $cpsform['sampleorderno'][$col]);
    $spreadsheet->getActiveSheet()->getStyle($BC)->applyFromArray($styleArray);
    $spreadsheet->getActiveSheet()->getStyle($BC)->getAlignment()->setWrapText(true);

    if ( $cpsform['alist'][$col]['a7'][0] == 'on') {
        $completeIcon = 'completed:已出货';
    } else {
        $completeIcon = '未出货';
    }
    $BC = "{$Brow}2:{$Crow}2";
    //$spreadsheet->getActiveSheet()->mergeCells($BC);
    $spreadsheet->getActiveSheet()->setCellValue($Crow.'2' , $completeIcon);
    $spreadsheet->getActiveSheet()->getStyle($BC)->applyFromArray($styleArray);
    $spreadsheet->getActiveSheet()->getStyle($BC)->getAlignment()->setWrapText(true);

    $BC = "{$Brow}3:{$Crow}3";
    //$spreadsheet->getActiveSheet()->mergeCells($BC);
    $spreadsheet->getActiveSheet()->setCellValue($Crow.'3' , $cpsform['alist'][$col]['a2'][0]);
    $spreadsheet->getActiveSheet()->getStyle($BC)->applyFromArray($styleArray);
    $spreadsheet->getActiveSheet()->getStyle($BC)->getAlignment()->setWrapText(true);

    $headertitle = array('Factory','Sample No.：','Job no.：','Style no：','Shipment date：','style type','weight (kg)');
    $headerrow = 10;
    foreach ($headertitle as $value){
        $spreadsheet->getActiveSheet()->setCellValue('A'.$headerrow , $value);
        $spreadsheet->getActiveSheet()->getStyle('A'.$headerrow)->applyFromArray($styleArray);
        $spreadsheet->getActiveSheet()->getStyle('A'.$headerrow)->getAlignment()->setWrapText(true);
        $headerrow++;
    }

    $BC = "{$Brow}10:{$Crow}10";
    $spreadsheet->getActiveSheet()->mergeCells($BC);
    $spreadsheet->getActiveSheet()->setCellValue($Brow.'10' , $cpsform['alist'][$col]['a1'][0]);
    $spreadsheet->getActiveSheet()->getStyle($BC)->applyFromArray($styleArray);
    $spreadsheet->getActiveSheet()->getStyle($BC)->getAlignment()->setWrapText(true);

    $BC = "{$Brow}11:{$Crow}11";
    $spreadsheet->getActiveSheet()->mergeCells($BC);
    $spreadsheet->getActiveSheet()->setCellValue($Brow.'11' , $cpsform['sampleno'][$col]);
    $spreadsheet->getActiveSheet()->getStyle($BC)->applyFromArray($styleArray);
    $spreadsheet->getActiveSheet()->getStyle($BC)->getAlignment()->setWrapText(true);

    $BC = "{$Brow}12:{$Crow}12";
    $spreadsheet->getActiveSheet()->mergeCells($BC);
    $spreadsheet->getActiveSheet()->setCellValue($Brow.'12' , $cpsform['jobno'][$col]);
    $spreadsheet->getActiveSheet()->getStyle($BC)->applyFromArray($styleArray);
    $spreadsheet->getActiveSheet()->getStyle($BC)->getAlignment()->setWrapText(true);

    $BC = "{$Brow}13:{$Crow}13";
    $spreadsheet->getActiveSheet()->mergeCells($BC);
    $spreadsheet->getActiveSheet()->setCellValue($Brow.'13' , $cpsform['styleno'][$col]);
    $spreadsheet->getActiveSheet()->getStyle($BC)->applyFromArray($styleArray);
    $spreadsheet->getActiveSheet()->getStyle($BC)->getAlignment()->setWrapText(true);

    $BC = "{$Brow}14:{$Crow}14";
    $spreadsheet->getActiveSheet()->mergeCells($BC);
    $spreadsheet->getActiveSheet()->setCellValue($Brow.'14' , $cpsform['shipmentdate'][$col]);
    $spreadsheet->getActiveSheet()->getStyle($BC)->applyFromArray($styleArray);
    $spreadsheet->getActiveSheet()->getStyle($BC)->getAlignment()->setWrapText(true);



    $spreadsheet->getActiveSheet()->setCellValue($Brow.'15' , $cpsform['alist'][$col]['a3'][0]);
    $spreadsheet->getActiveSheet()->getStyle($Brow.'15')->applyFromArray($styleArray);
    $spreadsheet->getActiveSheet()->getStyle($Brow.'15')->getAlignment()->setWrapText(true);

    $spreadsheet->getActiveSheet()->setCellValue($Brow.'16' , $cpsform['alist'][$col]['a5'][0]);
    $spreadsheet->getActiveSheet()->getStyle($Brow.'16')->applyFromArray($styleArray);
    $spreadsheet->getActiveSheet()->getStyle($Brow.'16')->getAlignment()->setWrapText(true);

    $BC = "{$Crow}15:{$Crow}16";
    $spreadsheet->getActiveSheet()->mergeCells($BC);
    $spreadsheet->getActiveSheet()->setCellValue($Crow.'15' , $cpsform['alist'][$col]['a4'][0]);
    $spreadsheet->getActiveSheet()->getStyle($BC)->applyFromArray($styleArray);
    $spreadsheet->getActiveSheet()->getStyle($BC)->getAlignment()->setWrapText(true);

}

/**
 * 头部
 */

/**
 * 图片模块
 */
for($col=0;$col< count($cpsform['id']);$col++) {
    $Brow = chr(66 + $col * 2);  //B
    $img = $cpsform['alist'][$col]['a6'][0];
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


    if ($haveimg) {
        preg_match('/.(jpg|gif|bmp|jpeg|png)/i', $img, $imgformat);
        $imgformat = $imgformat[1];
        switch ($imgformat) {
            case "jpg":
            case "jpeg":
                $img = imagecreatefromjpeg($img);
                break;
            case "bmp":
                $img = imagecreatefromwbmp($img);
                break;
            case "gif":
                $img = imagecreatefromgif($img);
                break;
            case "png":
                $img = imagecreatefrompng($img);
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

        //$drawing->setWidth($width>250 ? 250:$width);
        $drawing->setHeight($height > 130 ? 130 : $height);
//$drawing->setHeight(150);


        //$drawing->setCoordinates($cola.'2');
        $drawing->setCoordinates($Brow.'3');
        $drawing->setOffsetX(5);
        $drawing->setOffsetY(5);
        $drawing->setWorksheet($spreadsheet->getActiveSheet());
    }
}
/* 图片模块 */


/**
 *  pcs
 */
for($col=0;$col< count($cpsform['id']);$col++) {
    $Brow = chr(66 + $col * 2);  //B
    $Crow = chr(67 + $col * 2);  //C
    if ($cpsform['shipmentlist'][$col][0] > 0) {
        $thisrow = 4;
        for ($u = 0,$i=1;$u < $cpsform['shipmentlist'][$col][0];  $u++,$i++) {
            $pcstitle = array('UK', 'FR/UK', 'FR/HK', 'US', 'AUS', 'print');



            $spreadsheet->getActiveSheet()->setCellValue("A".$thisrow, $pcstitle[$u]);
            $spreadsheet->getActiveSheet()->getStyle("A" . $thisrow)->applyFromArray($styleArray);
            $spreadsheet->getActiveSheet()->getStyle("A" . $thisrow)->getAlignment()->setWrapText(true);

            $smb = isselect($cpsform['shipmentlist'][$col]['sma'.$i]).$cpsform['shipmentlist'][$col]['smb'.$i][0].'  '.$cpsform['shipmentlist'][$col]['smb'.$i][1].' '.$cpsform['shipmentlist'][$col]['smb'.$i][2];
            $spreadsheet->getActiveSheet()->setCellValue($Brow. $thisrow, $smb);
            $spreadsheet->getActiveSheet()->getStyle($Brow. $thisrow)->applyFromArray($styleArray);
            $spreadsheet->getActiveSheet()->getStyle($Brow. $thisrow)->getAlignment()->setWrapText(true);
            $smc = ' ' ;
            if($cpsform['shipmentlist'][$col]['smc'.$i] == 'on'){
                $smc = '  已出货' ;
            }
            $spreadsheet->getActiveSheet()->setCellValue($Crow . $thisrow, $cpsform['shipmentlist'][$col]['smb'.$i][3].$smc);
            $spreadsheet->getActiveSheet()->getStyle($Crow . $thisrow)->applyFromArray($styleArray);
            $spreadsheet->getActiveSheet()->getStyle($Crow . $thisrow)->getAlignment()->setWrapText(true);
            $thisrow++;
        }

    }
}
/**
 *  //pcs
 */

/**
 * remark
 */
for($col=0;$col< count($cpsform['id']);$col++) {
    $Brow = chr(66 + $col * 2);  //B
    $Crow = chr(67 + $col * 2);  //C

    if($cpsform['elist'][$col]['fromnume'] >0){
        $thisrow = 19;
        for ($u = 0,$i=1;$u < $cpsform['elist'][$col]['fromnume'];  $u++,$i++) {

            $spreadsheet->getActiveSheet()->setCellValue('A'.$thisrow , $cpsform['elist'][$col]['e1'][$u]);
            $spreadsheet->getActiveSheet()->getStyle('A'.$thisrow )->applyFromArray($styleArray);
            $spreadsheet->getActiveSheet()->getStyle('A'.$thisrow )->getAlignment()->setWrapText(true);

            $spreadsheet->getActiveSheet()->setCellValue($Brow.$thisrow , $cpsform['elist'][$col]['e2'][$u]);
            $spreadsheet->getActiveSheet()->getStyle($Brow.$thisrow)->applyFromArray($styleArray);
            $spreadsheet->getActiveSheet()->getStyle($Brow.$thisrow)->getAlignment()->setWrapText(true);

            $spreadsheet->getActiveSheet()->setCellValue($Crow.$thisrow , $cpsform['elist'][$col]['e3'][$u]);
            $spreadsheet->getActiveSheet()->getStyle($Crow.$thisrow)->applyFromArray($styleArray);
            $spreadsheet->getActiveSheet()->getStyle($Crow.$thisrow)->getAlignment()->setWrapText(true);
            $thisrow++;
        }
    }

}
/**
 *  //remark
 */



/**
 * Total Cost
 */
//$spreadsheet->getActiveSheet()->setCellValue("A16", $cpsform["elist"]["fixedval"]["fixedtitle"][2]);
//$spreadsheet->getActiveSheet()->getStyle("A16")->applyFromArray($styleArray);
//$spreadsheet->getActiveSheet()->setCellValue("A17", $cpsform["elist"]["fixedval"]["fixedtitle"][3]);
//$spreadsheet->getActiveSheet()->getStyle("A17")->applyFromArray($styleArray);
//$spreadsheet->getActiveSheet()->setCellValue("A18", $cpsform["elist"]["fixedval"]["fixedtitle"][4]);
//$spreadsheet->getActiveSheet()->getStyle("A18")->applyFromArray($styleArray);
//
//$fixa4 = getforexcate($cpsform["fixalist"]["fixa4"]).' '.$cpsform["elist"]["fixedval"]["fixedval"][2];
//$spreadsheet->getActiveSheet()->setCellValue("B16", $fixa4);
//$fixedval3 = $cpsform["elist"]["fixedval"]["fixedval"][3].' %';
//$spreadsheet->getActiveSheet()->setCellValue("B17", $fixedval3);
//$spreadsheet->getActiveSheet()->setCellValue("B18", $cpsform["elist"]["fixedval"]["fixedval"][4]);
//
//$spreadsheet->getActiveSheet()->getStyle("B16")->applyFromArray($styleArray);
//$spreadsheet->getActiveSheet()->getStyle("B17")->applyFromArray($styleArray);
//$spreadsheet->getActiveSheet()->getStyle("B18")->applyFromArray($styleArray);
//$spreadsheet->getActiveSheet()->getStyle("B16")->getAlignment()->setWrapText(true);
//$spreadsheet->getActiveSheet()->getStyle("B17")->getAlignment()->setWrapText(true);
//$spreadsheet->getActiveSheet()->getStyle("B18")->getAlignment()->setWrapText(true);
//
//$spreadsheet->getActiveSheet()->getStyle("C16")->applyFromArray($styleArray);
//$spreadsheet->getActiveSheet()->getStyle("C17")->applyFromArray($styleArray);
//$spreadsheet->getActiveSheet()->getStyle("C18")->applyFromArray($styleArray);

/**
 *  //Total Cost
 */



/**
 *  Total Trim Cost row+
 */
for($col=0;$col< count($cpsform['id']);$col++) {
    $Brow = chr(66 + $col * 2);  //B
    $Crow = chr(67 + $col * 2);  //C
    if ($cpsform["blist"][$col][0] > 0) {
        $titlearr = array('物料 & 特殊工序', 'Production Booking date', 'Colour standard received', 'Lab dips submitted', 'Lab dips approved', 'Base test report', 'Bulk cloth submitted', 'Bulk cloth approved', 'Bulk test  report approved', 'Bulk fabric  ready in factory', '上布方式', 'Care label', '1st Proto submitted', '1st Proto approved', '2nd Proto submitted', '2nd proto approved  (sealed to red seal)', 'Black seal  sample submitted', 'Black seal  sample approved', '1st off sample  submitted/approved', 'Shipment Sample', '上开工辦日期');

        $thisrow = 19;
        for ($u = 0,$i=1;$u < 21;  $u++,$i++) {

            if($col == 0){
                $spreadsheet->getActiveSheet()->insertNewRowBefore(17, 1);
            }


            $spreadsheet->getActiveSheet()->setCellValue('A'.$thisrow , $titlearr[$u]);
            $spreadsheet->getActiveSheet()->getStyle("A" . $thisrow)->getAlignment()->setWrapText(true);

            $thisrow++;
        }

        $thisrow = 19;
        for ($u = 0,$i=1;$u < 21;  $u++,$i++) {

            $spreadsheet->getActiveSheet()->setCellValue($Brow. $thisrow, $cpsform["blist"][$col]["b".$i][0]);
            $spreadsheet->getActiveSheet()->getStyle($Brow. $thisrow)->applyFromArray($styleArray);
            $spreadsheet->getActiveSheet()->getStyle($Brow. $thisrow)->getAlignment()->setWrapText(true);

            $spreadsheet->getActiveSheet()->setCellValue($Crow. $thisrow, $cpsform["clist"][$col]["c".$i][0]);
            $spreadsheet->getActiveSheet()->getStyle($Crow. $thisrow)->applyFromArray($styleArray);
            $spreadsheet->getActiveSheet()->getStyle($Crow. $thisrow)->getAlignment()->setWrapText(true);
            $thisrow++;
        }

    }




}
/**
 *  //Total Trim Cost row+
 */
/**
 * fa2alist
 */
for($col=0;$col< count($cpsform['id']);$col++) {
    $Brow = chr(66 + $col * 2);  //B
    $Crow = chr(67 + $col * 2);  //C

    if($cpsform['falist'][$col]['fa2alist'][0] >0){
        $fab2titlearr = array('最新紙樣資料(Merchandise)', '訂布資料(單位: Y/件)');
        $thisrow = 17;
        for ($u = 0,$i=1;$u < count($cpsform['falist'][$col]['fa2alist']['fa2a1']);  $u++,$i++) {

            $spreadsheet->getActiveSheet()->setCellValue('A'.$thisrow , $fab2titlearr[$u]);
            $spreadsheet->getActiveSheet()->getStyle('A'.$thisrow )->applyFromArray($styleArray);
            $spreadsheet->getActiveSheet()->getStyle('A'.$thisrow )->getAlignment()->setWrapText(true);

            $BC = "{$Brow}{$thisrow}:{$Crow}{$thisrow}";
            $spreadsheet->getActiveSheet()->mergeCells($BC);
            $spreadsheet->getActiveSheet()->setCellValue($Brow.$thisrow , $cpsform['falist'][$col]['fa2alist']['fa2a1'][$u]);
            $spreadsheet->getActiveSheet()->getStyle($BC)->applyFromArray($styleArray);
            $spreadsheet->getActiveSheet()->getStyle($BC)->getAlignment()->setWrapText(true);


            $thisrow++;
        }
    }

}

/**
 *  //fa2alist
 */
/**
 * shell
 */
for($col=0;$col< count($cpsform['id']);$col++) {
    $Brow = chr(66 + $col * 2);  //B
    $Crow = chr(67 + $col * 2);  //C
    if (isset($cpsform["falist"][$col]['falist']['fabrow']) && $cpsform["falist"][$col]['falist']['fabrow'] > 0) {

        $spreadsheet->getActiveSheet()->setCellValue('A19' , $cpsform["falist"][$col]['falist']['fabrow']);
        $spreadsheet->getActiveSheet()->getStyle("A19")->getAlignment()->setWrapText(true);

        $thisrow = 19;
        for ($u = 0,$i=1;$u < $cpsform["falist"][$col]['falist']['fabrow'];  $u++,$i++) {

//            if($col == 0){
//                $spreadsheet->getActiveSheet()->insertNewRowBefore(19, 1);
//            }


//            $spreadsheet->getActiveSheet()->setCellValue('A'.$thisrow , $cpsform["falist"][$col]['falist']['fabrow']);
//            $spreadsheet->getActiveSheet()->getStyle("A" . $thisrow)->getAlignment()->setWrapText(true);

//            $spreadsheet->getActiveSheet()->setCellValue($Brow. $thisrow, $cpsform["blist"][$col]["b".$i][0]);
//            $spreadsheet->getActiveSheet()->getStyle($Brow. $thisrow)->applyFromArray($styleArray);
//            $spreadsheet->getActiveSheet()->getStyle($Brow. $thisrow)->getAlignment()->setWrapText(true);
//
//            $spreadsheet->getActiveSheet()->setCellValue($Crow. $thisrow, $cpsform["clist"][$col]["c".$i][0]);
//            $spreadsheet->getActiveSheet()->getStyle($Crow. $thisrow)->applyFromArray($styleArray);
//            $spreadsheet->getActiveSheet()->getStyle($Crow. $thisrow)->getAlignment()->setWrapText(true);
            $thisrow++;
        }

//        $thisrow = 19;
//        for ($u = 0,$i=1;$u < 21;  $u++,$i++) {
//
//            $spreadsheet->getActiveSheet()->setCellValue($Brow. $thisrow, $cpsform["blist"][$col]["b".$i][0]);
//            $spreadsheet->getActiveSheet()->getStyle($Brow. $thisrow)->applyFromArray($styleArray);
//            $spreadsheet->getActiveSheet()->getStyle($Brow. $thisrow)->getAlignment()->setWrapText(true);
//
//            $spreadsheet->getActiveSheet()->setCellValue($Crow. $thisrow, $cpsform["clist"][$col]["c".$i][0]);
//            $spreadsheet->getActiveSheet()->getStyle($Crow. $thisrow)->applyFromArray($styleArray);
//            $spreadsheet->getActiveSheet()->getStyle($Crow. $thisrow)->getAlignment()->setWrapText(true);
//            $thisrow++;
//        }

    }

}

/**
 *  //shell
 */
/**
 *  Total Trim Cost row+
 */

//if($cpsform["clist"]['fromnume'] > 0){
//
//    for ($u = ($cpsform["clist"]['fromnume'] - 1);$u >= 0;$u-- ){
//        $thisrow = 14;
//        $spreadsheet->getActiveSheet()->insertNewRowBefore(14, 1);
//
//        $spreadsheet->getActiveSheet()->setCellValue("A14", $cpsform["clist"]["titlee"][$u]);
//        $spreadsheet->getActiveSheet()->getStyle("A".$thisrow)->applyFromArray($styleArray);
//        $spreadsheet->getActiveSheet()->getStyle("A".$thisrow)->getAlignment()->setWrapText(true);
//
//
//        $spreadsheet->getActiveSheet()->setCellValue("B".$thisrow, $cpsform["clist"]["c1"][$u]);
//        $spreadsheet->getActiveSheet()->getStyle("B".$thisrow)->applyFromArray($styleArray);
//        $spreadsheet->getActiveSheet()->getStyle("B".$thisrow)->getAlignment()->setWrapText(true);
//
//        $spreadsheet->getActiveSheet()->setCellValue("C".$thisrow, $cpsform["clist"]["c2"][$u]);
//        $spreadsheet->getActiveSheet()->getStyle("C".$thisrow)->applyFromArray($styleArray);
//        $spreadsheet->getActiveSheet()->getStyle("C".$thisrow)->getAlignment()->setWrapText(true);
//
//    }
//
//}

/**
 *  //Total Trim Cost row+
 */

/**
 *  主布
 */
//
//if($cpsform["alist"]["a10"] > 0){
//
//    for ($u = ($cpsform["alist"]["a10"] - 1);$u >= 0;$u-- ){
//        $thisrow = 11;
//        $spreadsheet->getActiveSheet()->insertNewRowBefore(11, 3);
//
//        $spreadsheet->getActiveSheet()->setCellValue("A11", $cpsform["alist"]["a11"][$u]);
//        $spreadsheet->getActiveSheet()->getStyle("A".$thisrow)->applyFromArray($styleArray);
//        $spreadsheet->getActiveSheet()->getStyle("A".$thisrow)->getAlignment()->setWrapText(true);
//
//
//        $spreadsheet->getActiveSheet()->setCellValue("B11", $cpsform["alist"]["a12"][$u]);
//        $spreadsheet->getActiveSheet()->getStyle("B".$thisrow)->applyFromArray($styleArray);
//        $spreadsheet->getActiveSheet()->getStyle("B".$thisrow)->getAlignment()->setWrapText(true);
//
//        $spreadsheet->getActiveSheet()->setCellValue("C11", $cpsform["alist"]["a13"][$u]);
//        $spreadsheet->getActiveSheet()->getStyle("C".$thisrow)->applyFromArray($styleArray);
//        $spreadsheet->getActiveSheet()->getStyle("C".$thisrow)->getAlignment()->setWrapText(true);
//
//        $thisrow = 12;
//        $spreadsheet->getActiveSheet()->setCellValue("A12", $cpsform["alist"]["a14"][$u]);
//        $spreadsheet->getActiveSheet()->getStyle("A".$thisrow)->applyFromArray($styleArray);
//        $spreadsheet->getActiveSheet()->getStyle("A".$thisrow)->getAlignment()->setWrapText(true);
//
//        if($cpsform["alist"]["a4"][$u] == 1){
//            $A4 = ' /y';
//        }else{
//            $A4 = ' /m';
//        }
//        $B12 = getforexcate($cpsform["alist"]["a15"][$u]).' '.$cpsform["alist"]["a16"][$u].$A4.' '.$cpsform["alist"]["a17"][$u];
//        $spreadsheet->getActiveSheet()->setCellValue("B".$thisrow, $B12);
//        $spreadsheet->getActiveSheet()->getStyle("B".$thisrow)->applyFromArray($styleArray);
//        $spreadsheet->getActiveSheet()->getStyle("B".$thisrow)->getAlignment()->setWrapText(true);
//
//        $spreadsheet->getActiveSheet()->setCellValue("C".$thisrow, $cpsform["alist"]["a18"][$u]);
//        $spreadsheet->getActiveSheet()->getStyle("C".$thisrow)->applyFromArray($styleArray);
//        $spreadsheet->getActiveSheet()->getStyle("C".$thisrow)->getAlignment()->setWrapText(true);
//
//        $thisrow = 13;
//        $spreadsheet->getActiveSheet()->setCellValue("A13", $cpsform["alist"]["a19"][$u]);
//        $spreadsheet->getActiveSheet()->getStyle("A".$thisrow)->applyFromArray($styleArray);
//        $spreadsheet->getActiveSheet()->getStyle("A".$thisrow)->getAlignment()->setWrapText(true);
//
//        if($cpsform["alist"]["a22"][$u] == 1){
//            $A22 = ' y/DZ';
//        }else{
//            $A22 = ' y/PC';
//        }
//        $B13 = $cpsform["alist"]["a20"][$u].' X  '.$cpsform["alist"]["a21"][$u].$A22;
//        $spreadsheet->getActiveSheet()->setCellValue("B".$thisrow, $B13);
//        $spreadsheet->getActiveSheet()->getStyle("B".$thisrow)->applyFromArray($styleArray);
//        $spreadsheet->getActiveSheet()->getStyle("B".$thisrow)->getAlignment()->setWrapText(true);
//
//
//        if($cpsform["alist"]["a24"][$u] == 1){
//            $A24 = ' y/PC';
//        }else{
//            $A24 = ' m/PC';
//        }
//        $spreadsheet->getActiveSheet()->setCellValue("C".$thisrow, $cpsform["alist"]["a23"][$u].$A24);
//        $spreadsheet->getActiveSheet()->getStyle("C".$thisrow)->applyFromArray($styleArray);
//        $spreadsheet->getActiveSheet()->getStyle("C".$thisrow)->getAlignment()->setWrapText(true);
//    }
//
//}

/**
 *  //主布
 */


// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$spreadsheet->setActiveSheetIndex(0);

//unset($_SESSION['cpsform'] ); //注销SESSION

//$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页
$output=  ($_GET['action'] == 'formdown' )? 1:0;
$nt = date("YmdHis",time()); //转换为日期。
$filenameout = 'CPS'.$nt.'.xlsx';
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