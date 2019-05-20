<?php
require_once 'aidenfunc.php';

$costp2 =  $_SESSION['costp2'];

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/*
 * 思路 先填固定行 后增加 可变行
 * 1
 */
//var_dump($costp2);
//$temno = $costp2["temno"];
//$titlearr = unserialize(gzuncompress(base64_decode($costp2["cctitle"])));
//print_r($titlearr);

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

$spreadsheet->getActiveSheet()->setTitle("sheet1");
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

function getforexcate($forex) {
    switch ($forex){
        case 1:
            $output = 'USD';
            break;
        case 2:
            $output = 'HKD';
            break;
        case 3:
            $output = 'RMB';
            break;
        case 4:
            $output = 'EUR';
            break;
        case 5:
            $output = 'JPY';
            break;

        default:
            $output = 'USD';
            break;
    }
    return $output;
}

function fabricname($cate) {
    switch ($cate){
        case '0':
            $output = 'Shell Fabric';
            break;
        case '1':
            $output = 'Lining';
            break;
        case '2':
            $output = 'Contrast';
            break;
        case '3':
            $output = 'Contrast 1';
            break;
        case '4':
            $output = 'Contrast 2';
            break;
        case '5':
            $output = 'Contrast 3';
            break;
        case '6':
            $output = 'Contrast 4';
            break;

        default:
            $output = 'Contrast';
            break;
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

$spreadsheet->getActiveSheet()->setCellValue("A1", 'CLIENT:');
$spreadsheet->getActiveSheet()->getStyle("A1")->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle("A1")->getAlignment()->setWrapText(true);

$spreadsheet->getActiveSheet()->setCellValue("A2", 'SO no.');
$spreadsheet->getActiveSheet()->getStyle("A2")->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle("A2")->getAlignment()->setWrapText(true);

$spreadsheet->getActiveSheet()->setCellValue("A3", 'Sketch');
$spreadsheet->getActiveSheet()->getStyle("A3:A9")->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle("A3:A9")->getAlignment()->setWrapText(true);

$spreadsheet->getActiveSheet()->setCellValue("A10", 'Style no.：');
$spreadsheet->getActiveSheet()->getStyle("A10")->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle("A10")->getAlignment()->setWrapText(true);

$spreadsheet->getActiveSheet()->getStyle("B3:B10")->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle("B10")->applyFromArray($styleArray);


$spreadsheet->getActiveSheet()->setCellValue("B1", $costp2["clientname"]);
$spreadsheet->getActiveSheet()->getStyle("B1")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle("B1")->getAlignment()->setWrapText(true);

$spreadsheet->getActiveSheet()->setCellValue("B2", $costp2['so']);
$spreadsheet->getActiveSheet()->getStyle("B2")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle("B2")->getAlignment()->setWrapText(true);

$spreadsheet->getActiveSheet()->setCellValue("C1", $costp2["alist"]["a1"]);
$spreadsheet->getActiveSheet()->getStyle("C1")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle("C1")->getAlignment()->setWrapText(true);

$spreadsheet->getActiveSheet()->mergeCells("C3:C9");
$spreadsheet->getActiveSheet()->setCellValue("C3", $costp2["alist"]["a3"]);
$spreadsheet->getActiveSheet()->getStyle("C3:C9")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle("C3:C9")->getAlignment()->setWrapText(true);


$spreadsheet->getActiveSheet()->setCellValue("B10", $costp2["styleno"]);
$spreadsheet->getActiveSheet()->getStyle("B10")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle("B10")->getAlignment()->setWrapText(true);


$spreadsheet->getActiveSheet()->getStyle("C2")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle("C10")->applyFromArray($styleArray1);


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

    //$drawing->setWidth($width>250 ? 250:$width);
    $drawing->setWidthAndHeight(300,130);  //设置图片最大宽度 高度
    //$drawing->setHeight($height>130 ? 130:$height);
//$drawing->setHeight(150);


    //$drawing->setCoordinates($cola.'2');
    $drawing->setCoordinates('B3');
    $drawing->setOffsetX(5);
    $drawing->setOffsetY(5);
    $drawing->setWorksheet($spreadsheet->getActiveSheet());
}
/* 图片模块 */

/**
 * FABRIC COST
 */
$spreadsheet->getActiveSheet()->setCellValue("A11", 'FABRIC COST');
$spreadsheet->getActiveSheet()->getStyle("A11")->applyFromArray($styleArray);


$fabricCost = getforexcate($costp2["fixalist"]["fixa1"]).'  '.$costp2["alist"]["a32"] .'%  ' .$costp2["alist"]["a31"];

//
//$a32 = '   '.getforexcate($costp2["fixalist"]["fixa1"]).' '.$costp2["alist"]["a32"];
//$radiob .=$a32;

$spreadsheet->getActiveSheet()->setCellValue("B11", $fabricCost);
////$spreadsheet->getActiveSheet()->setCellValue("B12", $radiob);
$spreadsheet->getActiveSheet()->getStyle("B11")->applyFromArray($styleArray1);
//$spreadsheet->getActiveSheet()->getStyle("B12")->applyFromArray($styleArray1);

$spreadsheet->getActiveSheet()->getStyle("C11")->applyFromArray($styleArray1);
//$spreadsheet->getActiveSheet()->getStyle("C12")->applyFromArray($styleArray1);




if($costp2["alist"]["a29"] == 0){    //如果 Fushing 为0 不打印

}else{
    $spreadsheet->getActiveSheet()->setCellValue("A12", 'Fushing');
    $spreadsheet->getActiveSheet()->setCellValue("B12", $costp2["alist"]["a29"]);
}

$spreadsheet->getActiveSheet()->getStyle("A12")->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->getStyle("B12")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle("B12")->getAlignment()->setWrapText(true);

//$spreadsheet->getActiveSheet()->setCellValue("C12", $costp2["alist"]["a30"]);
//$spreadsheet->getActiveSheet()->getStyle("C12")->applyFromArray($styleArray1);
//$spreadsheet->getActiveSheet()->getStyle("C12")->getAlignment()->setWrapText(true);
setCell($sheet,"C12",stripcslashes($costp2["alist"]["a30"]),$Size12borderscenter);
/**
 *  //FABRIC COST
 */

/**
 * Total Trim Cost
 */
$spreadsheet->getActiveSheet()->setCellValue("A13", $costp2["elist"]["fixedval"]["fixedtitle"][0]);
$spreadsheet->getActiveSheet()->getStyle("A13")->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->setCellValue("A14", $costp2["elist"]["fixedval"]["fixedtitle"][1]);
$spreadsheet->getActiveSheet()->getStyle("A14")->applyFromArray($styleArray);

$fixa2 = getforexcate($costp2["fixalist"]["fixa2"]).' '.$costp2["elist"]["fixedval"]["fixedval"][0];

$fixa3 = getforexcate($costp2["fixalist"]["fixa3"]).' '.$costp2["elist"]["fixedval"]["fixedval"][1];


$spreadsheet->getActiveSheet()->setCellValue("B13", $fixa2);
$spreadsheet->getActiveSheet()->setCellValue("B14", $fixa3);
$spreadsheet->getActiveSheet()->getStyle("B13")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle("B14")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle("B13")->getAlignment()->setWrapText(true);
$spreadsheet->getActiveSheet()->getStyle("B14")->getAlignment()->setWrapText(true);

$spreadsheet->getActiveSheet()->getStyle("C13")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle("C14")->applyFromArray($styleArray1);

/**
 *  //Total Trim Cost
 */

/**
 * Total Cost
 */
$spreadsheet->getActiveSheet()->setCellValue("A15", $costp2["elist"]["fixedval"]["fixedtitle"][2]);
$spreadsheet->getActiveSheet()->getStyle("A15")->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->setCellValue("A16", $costp2["elist"]["fixedval"]["fixedtitle"][3]);
$spreadsheet->getActiveSheet()->getStyle("A16")->applyFromArray($styleArray);
$spreadsheet->getActiveSheet()->setCellValue("A17", $costp2["elist"]["fixedval"]["fixedtitle"][4]);
$spreadsheet->getActiveSheet()->getStyle("A17")->applyFromArray($styleArray);

$fixa4 = getforexcate($costp2["fixalist"]["fixa4"]).' '.$costp2["elist"]["fixedval"]["fixedval"][2];
$spreadsheet->getActiveSheet()->setCellValue("B15", $fixa4);
$fixedval3 = $costp2["elist"]["fixedval"]["fixedval"][3].'%';
$spreadsheet->getActiveSheet()->setCellValue("B16", $fixedval3);

$B18 = getforexcate($costp2["fixalist"]["fixa4"]).' '.$costp2["elist"]["fixedval"]["fixedval"][4];
$spreadsheet->getActiveSheet()->setCellValue("B17", $B18);

$spreadsheet->getActiveSheet()->getStyle("B15")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle("B16")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle("B17")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle("B15")->getAlignment()->setWrapText(true);
$spreadsheet->getActiveSheet()->getStyle("B16")->getAlignment()->setWrapText(true);
$spreadsheet->getActiveSheet()->getStyle("B17")->getAlignment()->setWrapText(true);

$spreadsheet->getActiveSheet()->getStyle("C15")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle("C16")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->getStyle("C17")->applyFromArray($styleArray1);

/**
 *  //Total Cost
 */

/**
 * Unit Price
 */
if($costp2["glist"]["g1"] > 0){
    $spreadsheet->getActiveSheet()->setCellValue("A18", $costp2["elist"]["fixedval"]["fixedtitle"][5]);
    $spreadsheet->getActiveSheet()->getStyle("A18")->applyFromArray($styleArray);


    $b = 1;$v=0;

    foreach ($costp2["glist"]["g2"] as $value){
        if($b == $costp2["glist"]["g1"]){
            $g2 = '■  '.$value. '   '.$costp2["glist"]["g3"][$v] . '%';
        }else{
            $g2 = '   '.$value. '   '.$costp2["glist"]["g3"][$v] . '%';
        }
        $thisrow = 17 + $b ;
        $spreadsheet->getActiveSheet()->setCellValue("B".$thisrow , $g2);
        $spreadsheet->getActiveSheet()->getStyle("B".$thisrow)->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->getStyle("B".$thisrow)->getAlignment()->setWrapText(true);

        $spreadsheet->getActiveSheet()->getStyle("A".$thisrow)->applyFromArray($styleArray);

        $spreadsheet->getActiveSheet()->setCellValue("C".$thisrow , getforexcate($costp2["fixalist"]["fixa4"]).' '.$costp2["glist"]["g4"][$v]);
        $spreadsheet->getActiveSheet()->getStyle("C".$thisrow)->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->getStyle("C".$thisrow)->getAlignment()->setWrapText(true);
        $b++;$v++;
    }
    /**
     *  Final Price
     */
    $thisrow = 17 + $b ;
    $spreadsheet->getActiveSheet()->setCellValue("A".$thisrow, $costp2["elist"]["fixedval"]["fixedtitle"][6]);
    $spreadsheet->getActiveSheet()->getStyle("A".$thisrow)->applyFromArray($styleArray);

    $spreadsheet->getActiveSheet()->setCellValue("B".$thisrow, getforexcate($costp2["fixalist"]["fixa4"]).' '.$costp2["elist"]["fixedval"]["fixedval"][5]);
    $spreadsheet->getActiveSheet()->getStyle("B".$thisrow)->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->getStyle("B".$thisrow)->getAlignment()->setWrapText(true);

    $spreadsheet->getActiveSheet()->getStyle("C".$thisrow)->applyFromArray($styleArray1);
}


/**
 *  //Unit Price
 */

/**
 *  Total Trim Cost row+
 */

if($costp2["dlist"]['fromnumf'] > 0){

    for ($u = ($costp2["dlist"]['fromnumf'] - 1);$u >= 0;$u-- ){
        $thisrow = 15;
        $spreadsheet->getActiveSheet()->insertNewRowBefore(15, 1);

        $spreadsheet->getActiveSheet()->setCellValue("A15", $costp2["dlist"]["titlef"][$u]);
        $spreadsheet->getActiveSheet()->getStyle("A".$thisrow)->applyFromArray($styleArray);
        $spreadsheet->getActiveSheet()->getStyle("A".$thisrow)->getAlignment()->setWrapText(true);


        $spreadsheet->getActiveSheet()->setCellValue("B".$thisrow, $costp2["dlist"]["d1"][$u]);
        $spreadsheet->getActiveSheet()->getStyle("B".$thisrow)->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->getStyle("B".$thisrow)->getAlignment()->setWrapText(true);

        $spreadsheet->getActiveSheet()->setCellValue("C".$thisrow, $costp2["dlist"]["d2"][$u]);
        $spreadsheet->getActiveSheet()->getStyle("C".$thisrow)->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->getStyle("C".$thisrow)->getAlignment()->setWrapText(true);

    }

}

/**
 *  //Total Trim Cost row+
 */

/**
 *  Total Trim Cost row+
 */

if($costp2["clist"]['fromnume'] > 0){

//    for ($u = ($costp2["clist"]['fromnume'] - 1);$u >= 0;$u-- ){
    for ($u = ($costp2["clist"]['fromnume']);$u >= 0;$u-- ){
        $thisrow = 13;
        $spreadsheet->getActiveSheet()->insertNewRowBefore(13, 1);

        $spreadsheet->getActiveSheet()->setCellValue("A13", $costp2["clist"]["titlee"][$u]);
        $spreadsheet->getActiveSheet()->getStyle("A".$thisrow)->applyFromArray($styleArray);
        $spreadsheet->getActiveSheet()->getStyle("A".$thisrow)->getAlignment()->setWrapText(true);


        $spreadsheet->getActiveSheet()->setCellValue("B".$thisrow, $costp2["clist"]["c1"][$u]);
        $spreadsheet->getActiveSheet()->getStyle("B".$thisrow)->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->getStyle("B".$thisrow)->getAlignment()->setWrapText(true);

//        $spreadsheet->getActiveSheet()->setCellValue("C".$thisrow, $costp2["clist"]["c2"][$u]);
//        $spreadsheet->getActiveSheet()->getStyle("C".$thisrow)->applyFromArray($styleArray1);
//        $spreadsheet->getActiveSheet()->getStyle("C".$thisrow)->getAlignment()->setWrapText(true);
        setCell($sheet,"C13",stripcslashes($costp2["clist"]["c2"][$u]),$Size12borderscenter);
    }

}

/**
 *  //Total Trim Cost row+
 */

/**
 *  主布
 */

if($costp2["alist"]["a10"] > 0){

    for ($u = ($costp2["alist"]["a10"] - 1);$u >= 0;$u-- ){
        $thisrow = 11;
        $spreadsheet->getActiveSheet()->insertNewRowBefore(11, 3);

        //$spreadsheet->getActiveSheet()->setCellValue("A11", $costp2["alist"]["a11"][$u]);
        $spreadsheet->getActiveSheet()->setCellValue("A11", fabricname($costp2["alist"]["a25"][$u]));
        $spreadsheet->getActiveSheet()->getStyle("A".$thisrow)->applyFromArray($styleArray);
        $spreadsheet->getActiveSheet()->getStyle("A".$thisrow)->getAlignment()->setWrapText(true);


//        $spreadsheet->getActiveSheet()->setCellValue("B11", $costp2["alist"]["a12"][$u]);
//        $spreadsheet->getActiveSheet()->getStyle("B".$thisrow)->applyFromArray($styleArray1);
//        $spreadsheet->getActiveSheet()->getStyle("B".$thisrow)->getAlignment()->setWrapText(true);

        setCell($sheet,"B11",stripcslashes($costp2["alist"]["a12"][$u]),$Size12borderscenter);

//        $spreadsheet->getActiveSheet()->setCellValue("C11", $costp2["alist"]["a13"][$u]);
//        $spreadsheet->getActiveSheet()->getStyle("C".$thisrow)->applyFromArray($styleArray1);
//        $spreadsheet->getActiveSheet()->getStyle("C".$thisrow)->getAlignment()->setWrapText(true);
        setCell($sheet,"C11",stripcslashes($costp2["alist"]["a13"][$u]),$Size12borderscenter);

        $thisrow = 12;  //12
        $spreadsheet->getActiveSheet()->setCellValue("A12", $costp2["alist"]["a14"][$u]);
        $spreadsheet->getActiveSheet()->getStyle("A".$thisrow)->applyFromArray($styleArray);
        $spreadsheet->getActiveSheet()->getStyle("A".$thisrow)->getAlignment()->setWrapText(true);

        if($costp2["alist"]["a4"][$u] == 1){
            $A4 = ' /y';
        }else{
            $A4 = ' /m';
        }
        $B12 = getforexcate($costp2["alist"]["a15"][$u]).' '.$costp2["alist"]["a16"][$u].$A4.' '.$costp2["alist"]["a17"][$u];
        $spreadsheet->getActiveSheet()->setCellValue("B".$thisrow, $B12);
        $spreadsheet->getActiveSheet()->getStyle("B".$thisrow)->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->getStyle("B".$thisrow)->getAlignment()->setWrapText(true);

        $spreadsheet->getActiveSheet()->setCellValue("C".$thisrow, $costp2["alist"]["a18"][$u]);
        $spreadsheet->getActiveSheet()->getStyle("C".$thisrow)->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->getStyle("C".$thisrow)->getAlignment()->setWrapText(true);

        $thisrow = 13; //13
        $spreadsheet->getActiveSheet()->setCellValue("A13", $costp2["alist"]["a19"][$u]);
        $spreadsheet->getActiveSheet()->getStyle("A".$thisrow)->applyFromArray($styleArray);
        $spreadsheet->getActiveSheet()->getStyle("A".$thisrow)->getAlignment()->setWrapText(true);

        if($costp2["alist"]["a22"][$u] == 1){
            $A22 = ' y/DZ';
        }else{
            $A22 = ' y/PC';
        }
        $B13 = stripcslashes($costp2["alist"]["a20"][$u]).' X  '.$costp2["alist"]["a21"][$u].$A22;
        $spreadsheet->getActiveSheet()->setCellValue("B".$thisrow, $B13);
        $spreadsheet->getActiveSheet()->getStyle("B".$thisrow)->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->getStyle("B".$thisrow)->getAlignment()->setWrapText(true);

        if(($costp2["alist"]["a22"][$u] == 2) and ($costp2["alist"]["a24"][$u] == 1) ){

        }else{
            if($costp2["alist"]["a24"][$u] == 1){
                $A24 = ' y/PC';
            }else{
                $A24 = ' m/PC';
            }
//            $spreadsheet->getActiveSheet()->setCellValue("C".$thisrow, $costp2["alist"]["a23"][$u].$A24);
//            $spreadsheet->getActiveSheet()->getStyle("C".$thisrow)->applyFromArray($styleArray1);
//            $spreadsheet->getActiveSheet()->getStyle("C".$thisrow)->getAlignment()->setWrapText(true);
            setCell($sheet,"C".$thisrow,$costp2["alist"]["a23"][$u].$A24,$Size12borderscenter);
        }


    }

}

/**
 *  //主布
 */


// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$spreadsheet->setActiveSheetIndex(0);

unset($_SESSION['costp2'] ); //注销SESSION

$spreadsheet->getActiveSheet()->getPageSetup()
    ->setOrientation(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE);  //横放置
$spreadsheet->getActiveSheet()->getPageSetup()
    ->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);  //A4
$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页


$filenameout = "CostChart_{$costp2['shortname']}_";
outExcel($spreadsheet,$filenameout);

