<?php
session_start();
global $action;

$cpsall =  $_SESSION['cpsall'];
//var_dump($cpsp1);
//require '/home/soft/vendor/autoload.php';
//require '../vendor/autoload.php';
require '/home/pan/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;


    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/cpsall.xlsx');

    $sheet = $spreadsheet->getActiveSheet();


    $spreadsheet->getActiveSheet()->setTitle("sheet1");
//$sheet->setCellValue('A1', 'Hello World !');
    $spreadsheet->getDefaultStyle()->getFont()->setName('微软雅黑');
    $spreadsheet->getDefaultStyle()->getFont()->setSize(10);
//$spreadsheet->getActiveSheet()->getDefaultRowDimension()->setRowHeight(50);  //默认行高度
/*
    $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(5);//列宽度高度
    $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(13);//列宽度高度
    $spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(5);//列宽度高度
*/
$styleArray1 = [
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
        'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP,
        'wrapText' => true,
        'ShrinkToFit'=>true,
    ],
    'font' => [
        'Size' => '10',
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

$maxnum = $cpsall['cpsp1']["maxnum"];
$opensheetnum = 0; //当前打开的sheet数量

$prnum = $maxnum < 5 ? $maxnum : 4;

/**
 * $lt 列名
 * $lan 列數組序號
 * $lad 列數組數據 $lan + $prnum+1
 */
$spreadsheet->setActiveSheetIndex(0);  //設置當前活動表

$spreadsheet->getActiveSheet()->getRowDimension('2')->setRowHeight(90); //列高度
$spreadsheet->getActiveSheet()->getRowDimension('7')->setRowHeight(32); //列高度
$spreadsheet->getActiveSheet()->getRowDimension('13')->setRowHeight(32); //列高度
$spreadsheet->getActiveSheet()->getRowDimension('14')->setRowHeight(32); //列高度
$spreadsheet->getActiveSheet()->getRowDimension('15')->setRowHeight(32); //列高度
$spreadsheet->getActiveSheet()->getRowDimension('16')->setRowHeight(32); //列高度
$spreadsheet->getActiveSheet()->getRowDimension('17')->setRowHeight(32); //列高度




for($lt = 0, $lan = 0; $lt<=$prnum; $lt++){
    $lad = $lan + $maxnum + 1;
    //$col = chr(97 + $x);
    $cola = chr(66 + ($lt * 2)); //66 =B;
    $colb = chr(67 + ($lt * 2)); //66 =B;
    //echo '第一行'.$col.$i;
    $spreadsheet->getActiveSheet()->getColumnDimension($cola)->setWidth(15);  //列宽度
    $spreadsheet->getActiveSheet()->getColumnDimension($colb)->setWidth(20);  //列宽度

    $spreadsheet->getActiveSheet()->getStyle("{$cola}1:{$colb}1")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->setCellValue($cola.'1', $cpsall['cpsp1'][$lan][0]["cpsno"]);
//$spreadsheet->getActiveSheet()->setCellValue('C1', $cpsall['cpsp1']["maxnum"]);
    $spreadsheet->getActiveSheet()->mergeCells("{$cola}3:{$colb}3");
    $spreadsheet->getActiveSheet()->getStyle("{$cola}3:{$colb}3")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->setCellValue("{$cola}3", $cpsall['cpsp1'][$lan][0]["ftyno"]);

    $spreadsheet->getActiveSheet()->getStyle("{$cola}4:{$colb}4")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->mergeCells("{$cola}4:{$colb}4");
    $spreadsheet->getActiveSheet()->setCellValue("{$cola}4", $cpsall['cpsp1'][$lan][0]["jobno"]);

    $spreadsheet->getActiveSheet()->getStyle("{$cola}5:{$colb}5")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->mergeCells("{$cola}5:{$colb}5");
    $spreadsheet->getActiveSheet()->setCellValue("{$cola}5", $cpsall['cpsp1'][$lan][0]["styleno"]);

    $spreadsheet->getActiveSheet()->getStyle("{$cola}6:{$colb}6")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->mergeCells("{$cola}6:{$colb}6");
    $spreadsheet->getActiveSheet()->setCellValue("{$cola}6", $cpsall['cpsp1'][$lad][0]);

    $spreadsheet->getActiveSheet()->getStyle("{$cola}7:{$colb}7")->applyFromArray($styleArray1);

    $spreadsheet->getActiveSheet()->mergeCells("{$cola}7:{$colb}7");
    $spreadsheet->getActiveSheet()->setCellValue("{$cola}7", $cpsall['cpsp1'][$lad][1]);

    $spreadsheet->getActiveSheet()->getStyle("{$cola}8")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->setCellValue("{$cola}8", 'EUR');


    $spreadsheet->getActiveSheet()->getStyle("{$colb}8")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->setCellValue("{$colb}8", $cpsall['cpsp1'][$lad][2]);

    $spreadsheet->getActiveSheet()->getStyle("{$cola}9")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->setCellValue("{$cola}9", 'OTF');

    $spreadsheet->getActiveSheet()->getStyle("{$colb}9")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->setCellValue("{$colb}9", $cpsall['cpsp1'][$lad][3]);

    $spreadsheet->getActiveSheet()->getStyle("{$cola}2:{$colb}2")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->setCellValue("{$cola}2", $cpsall['cpsp1'][$lad][4]);

    /*加載圖片*/
    $img = $cpsall['cpsp1'][$lan][0]["remarkimg2"];
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


// Generate an image
//$gdImage = @imagecreatetruecolor($width, $height) or die('Cannot Initialize new GD image stream');
//$textColor = imagecolorallocate($gdImage, 255, 255, 255);
//imagestring($gdImage, 1, 5, 5,  'Created with PhpSpreadsheet', $textColor);

// Add a drawing to the worksheet
    $drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
    $drawing->setName($cpsall['cpsp1'][$lad][4]);
    $drawing->setDescription($cpsall['cpsp1'][$lad][4]);
//$drawing->setImageResource($gdImage);
    $drawing->setImageResource($img);
    $drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
    $drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
//$drawing->setHeight($width);

    $drawing->setHeight($height>100 ? 100:$height);
//$drawing->setWidth(180);
//$drawing->setHeight(150);
    $drawing->setCoordinates($colb.'2');
    $drawing->setOffsetX(5);
    $drawing->setOffsetY(5);
    $drawing->setWorksheet($spreadsheet->getActiveSheet());

    /*加載圖片*/
    $spreadsheet->getActiveSheet()->getStyle("{$cola}10")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->setCellValue("{$cola}10", $cpsall['cpsp1'][$lad][5]);
    $spreadsheet->getActiveSheet()->getStyle("{$colb}10")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->setCellValue("{$colb}10", $cpsall['cpsp1'][$lad][6]);

    $spreadsheet->getActiveSheet()->getStyle("{$cola}11")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->setCellValue("{$cola}11", $cpsall['cpsp1'][$lad][7]);
    $spreadsheet->getActiveSheet()->getStyle("{$colb}11")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->setCellValue("{$colb}11", $cpsall['cpsp1'][$lad][8]);

    $spreadsheet->getActiveSheet()->getStyle("{$cola}12")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->setCellValue("{$cola}12", 'TOTAL');
    $spreadsheet->getActiveSheet()->getStyle("{$colb}12")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->setCellValue("{$colb}12", $cpsall['cpsp1'][$lad][9]);

    $spreadsheet->getActiveSheet()->mergeCells("{$cola}13:{$colb}13");
    $spreadsheet->getActiveSheet()->getStyle("{$cola}13:{$colb}13")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->setCellValue("{$cola}13", $cpsall['cpsp1'][$lad][10]);

    $spreadsheet->getActiveSheet()->mergeCells("{$cola}14:{$colb}14");
    $spreadsheet->getActiveSheet()->getStyle("{$cola}14:{$colb}14")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->setCellValue("{$cola}14", $cpsall['cpsp1'][$lad][11]);

    $spreadsheet->getActiveSheet()->mergeCells("{$cola}15:{$colb}15");
    $spreadsheet->getActiveSheet()->getStyle("{$cola}15:{$colb}15")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->setCellValue("{$cola}15", $cpsall['cpsp1'][$lad][12]);

    $spreadsheet->getActiveSheet()->mergeCells("{$cola}16:{$colb}16");
    $spreadsheet->getActiveSheet()->getStyle("{$cola}16:{$colb}16")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->setCellValue("{$cola}16", $cpsall['cpsp1'][$lad][13]);

    $spreadsheet->getActiveSheet()->mergeCells("{$cola}17:{$colb}17");
    $spreadsheet->getActiveSheet()->getStyle("{$cola}17:{$colb}17")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->setCellValue("{$cola}17", $cpsall['cpsp1'][$lad][14]);

    $spreadsheet->getActiveSheet()->mergeCells("{$cola}18:{$colb}18");
    $spreadsheet->getActiveSheet()->getStyle("{$cola}18:{$colb}18")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->setCellValue("{$cola}18", $cpsall['cpsp1'][$lad][15]);

    $spreadsheet->getActiveSheet()->mergeCells("{$cola}19:{$colb}19");
    $spreadsheet->getActiveSheet()->getStyle("{$cola}19:{$colb}19")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->setCellValue("{$cola}19", $cpsall['cpsp1'][$lad][16]);

    for($j = 0 ,$z=17, $x = 20; $j < 15 ; $j++) {
        //$z=//數組值
        //$x = 20; //行數

        for ($i = 0; $i < 2; $i++) {
            $list = chr(66 + $i + ($lt * 2)); //66 =B;
            $spreadsheet->getActiveSheet()->getRowDimension($x)->setRowHeight(32); //列高度
            $spreadsheet->getActiveSheet()->getStyle($list.$x)->applyFromArray($styleArray1);
            $spreadsheet->getActiveSheet()->setCellValue($list.$x, $cpsall['cpsp1'][$lad][$z]);
            $z++;
        }
        $x++;
    }

    $lan++;
} //1st for


/**
 * 第二頁
 */
$prnum = $maxnum <= 9 ? $maxnum : 9;
if($maxnum > 4 ){

    $opensheetnum++;

    /**
     * $lt 列名
     * $lan 列數組序號 取數據
     * $lad 列數組數據 $lan + $prnum+1
     */
    $spreadsheet->setActiveSheetIndex($opensheetnum);  //設置當前活動表

    $styleArray1 = [
        'alignment' => [
            'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
            'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            'wrapText' => true,
            'ShrinkToFit'=>true,
        ],
        'font' => [
            'Size' => '10',
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

    $fabname=array("CPS"," ","Fty no.",'Job no.','Style no','Shipment date','Net Weight(Kg/pc)','PO & QTY(PCS)','PO & QTY(PCS)','PO & QTY(PCS)','PO & QTY(PCS)','PO & QTY(PCS)','Fabric composition','Lining','Trim fabric','最新紙樣資料(Merchandise)','訂布資料（單位 ：Y/件）','開裁用料（單位 ：Y/件）','物料','特殊工序','Production Booking date','Colour standard received','Lab dips submitted','Lab dips approved','D.farbic test report','Bulk cloth submitted','Bulk cloth approved','Bulk test report approved','Care label','Blue Seal Approved','Green Seal Submitted','Green Seal Approved','White Seal Approved','REMARK');
    $fabcou = count($fabname);
    for ($j=0,$row = 1;$j<$fabcou;$j++,$row++) {
        $spreadsheet->getActiveSheet()->setCellValue('A'.$row, $fabname[$j]);
        $spreadsheet->getActiveSheet()->getStyle('A'.$row)->applyFromArray($styleArray1);
    }
    $spreadsheet->getActiveSheet()->mergeCells("A8:A12");
    $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(20);  //列宽度

    $spreadsheet->getActiveSheet()->getRowDimension('2')->setRowHeight(90); //列高度
    $spreadsheet->getActiveSheet()->getRowDimension('7')->setRowHeight(32); //列高度
    $spreadsheet->getActiveSheet()->getRowDimension('13')->setRowHeight(32); //列高度
    $spreadsheet->getActiveSheet()->getRowDimension('14')->setRowHeight(32); //列高度
    $spreadsheet->getActiveSheet()->getRowDimension('15')->setRowHeight(32); //列高度
    $spreadsheet->getActiveSheet()->getRowDimension('16')->setRowHeight(32); //列高度
    $spreadsheet->getActiveSheet()->getRowDimension('17')->setRowHeight(32); //列高度




    for($lt = 0, $lan = 5,$time2 = 5; $time2<=$prnum; $lt++,$time2++){
        $lad = $lan + $maxnum + 1;//
        //$col = chr(97 + $x);
        $cola = chr(66 + ($lt * 2)); //66 =B;
        $colb = chr(67 + ($lt * 2)); //66 =B;
        //echo '第一行'.$col.$i;
        $spreadsheet->getActiveSheet()->getColumnDimension($cola)->setWidth(15);  //列宽度
        $spreadsheet->getActiveSheet()->getColumnDimension($colb)->setWidth(20);  //列宽度

        $spreadsheet->getActiveSheet()->getStyle("{$cola}1:{$colb}1")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue($cola.'1', $cpsall['cpsp1'][$lan][0]["cpsno"]);
//$spreadsheet->getActiveSheet()->setCellValue('C1', $cpsall['cpsp1']["maxnum"]);
        $spreadsheet->getActiveSheet()->mergeCells("{$cola}3:{$colb}3");
        $spreadsheet->getActiveSheet()->getStyle("{$cola}3:{$colb}3")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}3", $cpsall['cpsp1'][$lan][0]["ftyno"]);

        $spreadsheet->getActiveSheet()->getStyle("{$cola}4:{$colb}4")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->mergeCells("{$cola}4:{$colb}4");
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}4", $cpsall['cpsp1'][$lan][0]["jobno"]);

        $spreadsheet->getActiveSheet()->getStyle("{$cola}5:{$colb}5")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->mergeCells("{$cola}5:{$colb}5");
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}5", $cpsall['cpsp1'][$lan][0]["styleno"]);

        $spreadsheet->getActiveSheet()->getStyle("{$cola}6:{$colb}6")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->mergeCells("{$cola}6:{$colb}6");
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}6", $cpsall['cpsp1'][$lad][0]);

        $spreadsheet->getActiveSheet()->getStyle("{$cola}7:{$colb}7")->applyFromArray($styleArray1);

        $spreadsheet->getActiveSheet()->mergeCells("{$cola}7:{$colb}7");
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}7", $cpsall['cpsp1'][$lad][1]);

        $spreadsheet->getActiveSheet()->getStyle("{$cola}8")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}8", 'EUR');


        $spreadsheet->getActiveSheet()->getStyle("{$colb}8")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue("{$colb}8", $cpsall['cpsp1'][$lad][2]);

        $spreadsheet->getActiveSheet()->getStyle("{$cola}9")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}9", 'OTF');

        $spreadsheet->getActiveSheet()->getStyle("{$colb}9")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue("{$colb}9", $cpsall['cpsp1'][$lad][3]);

        $spreadsheet->getActiveSheet()->getStyle("{$cola}2:{$colb}2")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}2", $cpsall['cpsp1'][$lad][4]);

        /*加載圖片*/
        $img = $cpsall['cpsp1'][$lan][0]["remarkimg2"];
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


// Generate an image
//$gdImage = @imagecreatetruecolor($width, $height) or die('Cannot Initialize new GD image stream');
//$textColor = imagecolorallocate($gdImage, 255, 255, 255);
//imagestring($gdImage, 1, 5, 5,  'Created with PhpSpreadsheet', $textColor);

// Add a drawing to the worksheet
        $drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
        $drawing->setName($cpsall['cpsp1'][$lad][4]);
        $drawing->setDescription($cpsall['cpsp1'][$lad][4]);
//$drawing->setImageResource($gdImage);
        $drawing->setImageResource($img);
        $drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
        $drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
//$drawing->setHeight($width);
        $drawing->setWidth($width>150 ? 150:$width);

        //$drawing->setHeight($height>100 ? 100:$height);
//$drawing->setWidth(180);
//$drawing->setHeight(150);
        $drawing->setCoordinates($colb.'2');
        $drawing->setOffsetX(5);
        $drawing->setOffsetY(5);
        $drawing->setWorksheet($spreadsheet->getActiveSheet());

        /*加載圖片*/


        $spreadsheet->getActiveSheet()->getStyle("{$cola}10")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}10", $cpsall['cpsp1'][$lad][5]);
        $spreadsheet->getActiveSheet()->getStyle("{$colb}10")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue("{$colb}10", $cpsall['cpsp1'][$lad][6]);

        $spreadsheet->getActiveSheet()->getStyle("{$cola}11")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}11", $cpsall['cpsp1'][$lad][7]);
        $spreadsheet->getActiveSheet()->getStyle("{$colb}11")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue("{$colb}11", $cpsall['cpsp1'][$lad][8]);

        $spreadsheet->getActiveSheet()->getStyle("{$cola}12")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}12", 'TOTAL');
        $spreadsheet->getActiveSheet()->getStyle("{$colb}12")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue("{$colb}12", $cpsall['cpsp1'][$lad][9]);

        $spreadsheet->getActiveSheet()->mergeCells("{$cola}13:{$colb}13");
        $spreadsheet->getActiveSheet()->getStyle("{$cola}13:{$colb}13")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}13", $cpsall['cpsp1'][$lad][10]);

        $spreadsheet->getActiveSheet()->mergeCells("{$cola}14:{$colb}14");
        $spreadsheet->getActiveSheet()->getStyle("{$cola}14:{$colb}14")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}14", $cpsall['cpsp1'][$lad][11]);

        $spreadsheet->getActiveSheet()->mergeCells("{$cola}15:{$colb}15");
        $spreadsheet->getActiveSheet()->getStyle("{$cola}15:{$colb}15")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}15", $cpsall['cpsp1'][$lad][12]);

        $spreadsheet->getActiveSheet()->mergeCells("{$cola}16:{$colb}16");
        $spreadsheet->getActiveSheet()->getStyle("{$cola}16:{$colb}16")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}16", $cpsall['cpsp1'][$lad][13]);

        $spreadsheet->getActiveSheet()->mergeCells("{$cola}17:{$colb}17");
        $spreadsheet->getActiveSheet()->getStyle("{$cola}17:{$colb}17")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}17", $cpsall['cpsp1'][$lad][14]);

        $spreadsheet->getActiveSheet()->mergeCells("{$cola}18:{$colb}18");
        $spreadsheet->getActiveSheet()->getStyle("{$cola}18:{$colb}18")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}18", $cpsall['cpsp1'][$lad][15]);

        $spreadsheet->getActiveSheet()->mergeCells("{$cola}19:{$colb}19");
        $spreadsheet->getActiveSheet()->getStyle("{$cola}19:{$colb}19")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}19", $cpsall['cpsp1'][$lad][16]);

        for($j = 0 ,$z=17, $x = 20; $j < 15 ; $j++) {
            //$z=//數組值
            //$x = 20; //行數

            for ($i = 0; $i < 2; $i++) {
                $list = chr(66 + $i + ($lt * 2)); //66 =B;
                $spreadsheet->getActiveSheet()->getRowDimension($x)->setRowHeight(32); //列高度
                $spreadsheet->getActiveSheet()->getStyle($list.$x)->applyFromArray($styleArray1);
                $spreadsheet->getActiveSheet()->setCellValue($list.$x, $cpsall['cpsp1'][$lad][$z]);
                $z++;
            }
            $x++;
        }

        $lan++;
    } //1st for

}
$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

/**
 * 第三頁
 */
$prnum = $maxnum <=14 ? $maxnum : 14;
if($maxnum > 9 ){



    $opensheetnum++;
    /**
     * $lt 列名
     * $lan 列數組序號 取數據
     * $lad 列數組數據 $lan + $prnum+1
     */

    $spreadsheet->setActiveSheetIndex($opensheetnum);  //設置當前活動表

    $styleArray1 = [
        'alignment' => [
            'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
            'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            'wrapText' => true,
            'ShrinkToFit'=>true,
        ],
        'font' => [
            'Size' => '10',
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

    $fabname=array("CPS"," ","Fty no.",'Job no.','Style no','Shipment date','Net Weight(Kg/pc)','PO & QTY(PCS)','PO & QTY(PCS)','PO & QTY(PCS)','PO & QTY(PCS)','PO & QTY(PCS)','Fabric composition','Lining','Trim fabric','最新紙樣資料(Merchandise)','訂布資料（單位 ：Y/件）','開裁用料（單位 ：Y/件）','物料','特殊工序','Production Booking date','Colour standard received','Lab dips submitted','Lab dips approved','D.farbic test report','Bulk cloth submitted','Bulk cloth approved','Bulk test report approved','Care label','Blue Seal Approved','Green Seal Submitted','Green Seal Approved','White Seal Approved','REMARK');
    $fabcou = count($fabname);
    for ($j=0,$row = 1;$j<$fabcou;$j++,$row++) {
        $spreadsheet->getActiveSheet()->setCellValue('A'.$row, $fabname[$j]);
        $spreadsheet->getActiveSheet()->getStyle('A'.$row)->applyFromArray($styleArray1);
    }
    $spreadsheet->getActiveSheet()->mergeCells("A8:A12");
    $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(20);  //列宽度



    $spreadsheet->getActiveSheet()->getRowDimension('2')->setRowHeight(90); //列高度
    $spreadsheet->getActiveSheet()->getRowDimension('7')->setRowHeight(32); //列高度
    $spreadsheet->getActiveSheet()->getRowDimension('13')->setRowHeight(32); //列高度
    $spreadsheet->getActiveSheet()->getRowDimension('14')->setRowHeight(32); //列高度
    $spreadsheet->getActiveSheet()->getRowDimension('15')->setRowHeight(32); //列高度
    $spreadsheet->getActiveSheet()->getRowDimension('16')->setRowHeight(32); //列高度
    $spreadsheet->getActiveSheet()->getRowDimension('17')->setRowHeight(32); //列高度




    for($lt = 0, $lan = 10,$time3 = 10; $time3<=$prnum; $lt++,$time3++){
        $lad = $lan + $maxnum + 1;//
        //$col = chr(97 + $x);
        $cola = chr(66 + ($lt * 2)); //66 =B;
        $colb = chr(67 + ($lt * 2)); //66 =B;
        //echo '第一行'.$col.$i;
        $spreadsheet->getActiveSheet()->getColumnDimension($cola)->setWidth(15);  //列宽度
        $spreadsheet->getActiveSheet()->getColumnDimension($colb)->setWidth(20);  //列宽度

        $spreadsheet->getActiveSheet()->getStyle("{$cola}1:{$colb}1")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue($cola.'1', $cpsall['cpsp1'][$lan][0]["cpsno"]);
//$spreadsheet->getActiveSheet()->setCellValue('C1', $cpsall['cpsp1']["maxnum"]);
        $spreadsheet->getActiveSheet()->mergeCells("{$cola}3:{$colb}3");
        $spreadsheet->getActiveSheet()->getStyle("{$cola}3:{$colb}3")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}3", $cpsall['cpsp1'][$lan][0]["ftyno"]);

        $spreadsheet->getActiveSheet()->getStyle("{$cola}4:{$colb}4")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->mergeCells("{$cola}4:{$colb}4");
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}4", $cpsall['cpsp1'][$lan][0]["jobno"]);

        $spreadsheet->getActiveSheet()->getStyle("{$cola}5:{$colb}5")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->mergeCells("{$cola}5:{$colb}5");
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}5", $cpsall['cpsp1'][$lan][0]["styleno"]);

        $spreadsheet->getActiveSheet()->getStyle("{$cola}6:{$colb}6")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->mergeCells("{$cola}6:{$colb}6");
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}6", $cpsall['cpsp1'][$lad][0]);

        $spreadsheet->getActiveSheet()->getStyle("{$cola}7:{$colb}7")->applyFromArray($styleArray1);

        $spreadsheet->getActiveSheet()->mergeCells("{$cola}7:{$colb}7");
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}7", $cpsall['cpsp1'][$lad][1]);

        $spreadsheet->getActiveSheet()->getStyle("{$cola}8")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}8", 'EUR');


        $spreadsheet->getActiveSheet()->getStyle("{$colb}8")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue("{$colb}8", $cpsall['cpsp1'][$lad][2]);

        $spreadsheet->getActiveSheet()->getStyle("{$cola}9")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}9", 'OTF');

        $spreadsheet->getActiveSheet()->getStyle("{$colb}9")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue("{$colb}9", $cpsall['cpsp1'][$lad][3]);

        $spreadsheet->getActiveSheet()->getStyle("{$cola}2:{$colb}2")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}2", $cpsall['cpsp1'][$lad][4]);

        /*加載圖片*/
$img = $cpsall['cpsp1'][$lan][0]["remarkimg2"];
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


// Generate an image
//$gdImage = @imagecreatetruecolor($width, $height) or die('Cannot Initialize new GD image stream');
//$textColor = imagecolorallocate($gdImage, 255, 255, 255);
//imagestring($gdImage, 1, 5, 5,  'Created with PhpSpreadsheet', $textColor);

// Add a drawing to the worksheet
$drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
$drawing->setName($cpsall['cpsp1'][$lad][4]);
$drawing->setDescription($cpsall['cpsp1'][$lad][4]);
//$drawing->setImageResource($gdImage);
$drawing->setImageResource($img);
$drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
$drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
//$drawing->setHeight($width);
$drawing->setWidth($width>100 ? 100:$width);
//$drawing->setHeight($height>100 ? 100:$height);
//$drawing->setWidth(180);
//$drawing->setHeight(150);
$drawing->setCoordinates($colb.'2');
$drawing->setOffsetX(5);
$drawing->setOffsetY(5);
$drawing->setWorksheet($spreadsheet->getActiveSheet());

/*加載圖片*/
$spreadsheet->getActiveSheet()->getStyle("{$cola}10")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->setCellValue("{$cola}10", $cpsall['cpsp1'][$lad][5]);
$spreadsheet->getActiveSheet()->getStyle("{$colb}10")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->setCellValue("{$colb}10", $cpsall['cpsp1'][$lad][6]);

$spreadsheet->getActiveSheet()->getStyle("{$cola}11")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->setCellValue("{$cola}11", $cpsall['cpsp1'][$lad][7]);
$spreadsheet->getActiveSheet()->getStyle("{$colb}11")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->setCellValue("{$colb}11", $cpsall['cpsp1'][$lad][8]);

$spreadsheet->getActiveSheet()->getStyle("{$cola}12")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->setCellValue("{$cola}12", 'TOTAL');
$spreadsheet->getActiveSheet()->getStyle("{$colb}12")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->setCellValue("{$colb}12", $cpsall['cpsp1'][$lad][9]);

$spreadsheet->getActiveSheet()->mergeCells("{$cola}13:{$colb}13");
$spreadsheet->getActiveSheet()->getStyle("{$cola}13:{$colb}13")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->setCellValue("{$cola}13", $cpsall['cpsp1'][$lad][10]);

$spreadsheet->getActiveSheet()->mergeCells("{$cola}14:{$colb}14");
$spreadsheet->getActiveSheet()->getStyle("{$cola}14:{$colb}14")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->setCellValue("{$cola}14", $cpsall['cpsp1'][$lad][11]);

$spreadsheet->getActiveSheet()->mergeCells("{$cola}15:{$colb}15");
$spreadsheet->getActiveSheet()->getStyle("{$cola}15:{$colb}15")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->setCellValue("{$cola}15", $cpsall['cpsp1'][$lad][12]);

$spreadsheet->getActiveSheet()->mergeCells("{$cola}16:{$colb}16");
$spreadsheet->getActiveSheet()->getStyle("{$cola}16:{$colb}16")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->setCellValue("{$cola}16", $cpsall['cpsp1'][$lad][13]);

$spreadsheet->getActiveSheet()->mergeCells("{$cola}17:{$colb}17");
$spreadsheet->getActiveSheet()->getStyle("{$cola}17:{$colb}17")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->setCellValue("{$cola}17", $cpsall['cpsp1'][$lad][14]);

$spreadsheet->getActiveSheet()->mergeCells("{$cola}18:{$colb}18");
$spreadsheet->getActiveSheet()->getStyle("{$cola}18:{$colb}18")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->setCellValue("{$cola}18", $cpsall['cpsp1'][$lad][15]);

$spreadsheet->getActiveSheet()->mergeCells("{$cola}19:{$colb}19");
$spreadsheet->getActiveSheet()->getStyle("{$cola}19:{$colb}19")->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->setCellValue("{$cola}19", $cpsall['cpsp1'][$lad][16]);

for($j = 0 ,$z=17, $x = 20; $j < 15 ; $j++) {
    //$z=//數組值
    //$x = 20; //行數

    for ($i = 0; $i < 2; $i++) {
        $list = chr(66 + $i + ($lt * 2)); //66 =B;
        $spreadsheet->getActiveSheet()->getRowDimension($x)->setRowHeight(32); //列高度
        $spreadsheet->getActiveSheet()->getStyle($list.$x)->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue($list.$x, $cpsall['cpsp1'][$lad][$z]);
        $z++;
    }
    $x++;
}

$lan++;
} //1st for

};



$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页


/**
 * CPSP2  CPS第二份
 */
$opensheetnum++;
$spreadsheet->setActiveSheetIndex($opensheetnum);  //設置當前活動表

$sheet = $spreadsheet->getActiveSheet();


$spreadsheet->getDefaultStyle()->getFont()->setName('微软雅黑');
$spreadsheet->getDefaultStyle()->getFont()->setSize(10);

$fabname=array("CPS","Sketch：","Job no.",'Style no','Shipment date','Net Weight(g)','Fabric composition','Lining','Trim fabric','物料','訂布用料','最新用料（Y/件）','特殊工序','Lab dips submitted','Lab dips approved','Bulk cloth approved','Bulk test report approved','Bulk fabric ready in factory','上布方式','Care label','Blue seal approved','Green Seal approved','White seal lastest submit date','Remarks','QC & M.Detector report','SHORT SHIPPED');
$fabcou = count($fabname);
for ($j=0,$row = 1;$j<$fabcou;$j++,$row++) {
    $spreadsheet->getActiveSheet()->setCellValue('A'.$row, $fabname[$j]);
    $spreadsheet->getActiveSheet()->getStyle('A'.$row)->applyFromArray($styleArray1);
}

$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(20);  //列宽度

$styleArray1 = [
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
        'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP,
        'wrapText' => true,
        'ShrinkToFit'=>true,
    ],
    'font' => [
        'Size' => '10',
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

$maxnum = $cpsall['cpsp2']["maxnum"];


$prnum = $maxnum < 5 ? $maxnum : 4;

/**
 * $lt 列名
 * $lan 列數組序號
 * $lad 列數組數據 $lan + $prnum+1
 */


$spreadsheet->getActiveSheet()->getRowDimension('2')->setRowHeight(100); //列高度
$spreadsheet->getActiveSheet()->getRowDimension('7')->setRowHeight(32); //列高度
$spreadsheet->getActiveSheet()->getRowDimension('13')->setRowHeight(32); //列高度
$spreadsheet->getActiveSheet()->getRowDimension('14')->setRowHeight(32); //列高度
$spreadsheet->getActiveSheet()->getRowDimension('15')->setRowHeight(32); //列高度
$spreadsheet->getActiveSheet()->getRowDimension('16')->setRowHeight(32); //列高度
$spreadsheet->getActiveSheet()->getRowDimension('17')->setRowHeight(32); //列高度




for($lt = 0, $lan = 0; $lt<=$prnum; $lt++){
    $lad = $lan + $maxnum + 1;
    //$col = chr(97 + $x);
    $cola = chr(66 + ($lt * 2)); //66 =B;
    $colb = chr(67 + ($lt * 2)); //66 =B;
    //echo '第一行'.$col.$i;
    $spreadsheet->getActiveSheet()->getColumnDimension($cola)->setWidth(20);  //列宽度
    $spreadsheet->getActiveSheet()->getColumnDimension($colb)->setWidth(20);  //列宽度

    $spreadsheet->getActiveSheet()->getStyle("{$cola}1:{$colb}1")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->setCellValue($cola.'1', $cpsall['cpsp2'][$lan][0]["cpsno"]);

    $spreadsheet->getActiveSheet()->getStyle("{$colb}2")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->setCellValue("{$colb}2", $cpsall['cpsp2'][$lad][4]);
    /*
//$spreadsheet->getActiveSheet()->setCellValue('C1', $cpsall['cpsp2']["maxnum"]);
    $spreadsheet->getActiveSheet()->mergeCells("{$cola}3:{$colb}3");
    $spreadsheet->getActiveSheet()->getStyle("{$cola}3:{$colb}3")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->setCellValue("{$cola}3", $cpsall['cpsp2'][$lan][0]["ftyno"]);
*/
    $spreadsheet->getActiveSheet()->getStyle("{$cola}3:{$colb}3")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->mergeCells("{$cola}3:{$colb}3");
    $spreadsheet->getActiveSheet()->setCellValue("{$cola}3", $cpsall['cpsp2'][$lan][0]["jobno"]);

    $spreadsheet->getActiveSheet()->getStyle("{$cola}4:{$colb}4")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->mergeCells("{$cola}4:{$colb}4");
    $spreadsheet->getActiveSheet()->setCellValue("{$cola}4", $cpsall['cpsp2'][$lan][0]["styleno"]);



    for($listtop = 0,$lttime1 = 0,$listta = 5;$lttime1<=9;$listtop++,$listta++,$lttime1++){
        if($lttime1 == 4)
        {$listta = 8;
            continue;};
        $spreadsheet->getActiveSheet()->getStyle("{$cola}{$listta}:{$colb}{$listta}")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->mergeCells("{$cola}{$listta}:{$colb}{$listta}");
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}{$listta}", $cpsall['cpsp2'][$lad][$listtop]);
    }


    /*加載圖片*/
    $img = $cpsall['cpsp2'][$lan][0]["remarkimg2"];
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


// Generate an image
//$gdImage = @imagecreatetruecolor($width, $height) or die('Cannot Initialize new GD image stream');
//$textColor = imagecolorallocate($gdImage, 255, 255, 255);
//imagestring($gdImage, 1, 5, 5,  'Created with PhpSpreadsheet', $textColor);

// Add a drawing to the worksheet
    $drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
    $drawing->setName($cpsall['cpsp2'][$lad][4]);
    $drawing->setDescription($cpsall['cpsp2'][$lad][4]);
//$drawing->setImageResource($gdImage);
    $drawing->setImageResource($img);
    $drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
    $drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
//$drawing->setHeight($width);

    $drawing->setWidth($width>120 ? 120:$width);
//$drawing->setWidth(180);
//$drawing->setHeight(150);
    $drawing->setCoordinates($cola.'2');
    $drawing->setOffsetX(5);
    $drawing->setOffsetY(5);
    $drawing->setWorksheet($spreadsheet->getActiveSheet());

    /*加載圖片*/

    for($j = 0 ,$z=10, $x = 14; $j < 13 ; $j++) {
        //$z=//數組值
        //$x = 20; //行數

        for ($i = 0; $i < 2; $i++) {
            $list = chr(66 + $i + ($lt * 2)); //66 =B;
            $spreadsheet->getActiveSheet()->getRowDimension($x)->setRowHeight(32); //列高度
            $spreadsheet->getActiveSheet()->getStyle($list.$x)->applyFromArray($styleArray1);
            $spreadsheet->getActiveSheet()->setCellValue($list.$x,$cpsall['cpsp2'][$lad][$z] );
            $z++;
        }
        $x++;
    }

    $lan++;
} //1st for
$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

/**
 *  cps2 第二頁
 */
$prnum = $maxnum <= 9 ? $maxnum : 9;
if($maxnum > 4 ){
    $opensheetnum++;
    $spreadsheet->setActiveSheetIndex($opensheetnum);  //設置當前活動表
    /**
     * $lt 列名
     * $lan 列數組序號 取數據
     * $lad 列數組數據 $lan + $prnum+1
     */
    /*A列名称*/
    $fabname=array("CPS","Sketch：","Job no.",'Style no','Shipment date','Net Weight(g)','Fabric composition','Lining','Trim fabric','物料','訂布用料','最新用料（Y/件）','特殊工序','Lab dips submitted','Lab dips approved','Bulk cloth approved','Bulk test report approved','Bulk fabric ready in factory','上布方式','Care label','Blue seal approved','Green Seal approved','White seal lastest submit date','Remarks','QC & M.Detector report','SHORT SHIPPED');
    $fabcou = count($fabname);
    for ($j=0,$row = 1;$j<$fabcou;$j++,$row++) {
        $spreadsheet->getActiveSheet()->setCellValue('A'.$row, $fabname[$j]);
        $spreadsheet->getActiveSheet()->getStyle('A'.$row)->applyFromArray($styleArray1);
    }

    $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(20);  //列宽度
    /*A列名称*/


    for($lt = 0, $lan = 5,$time2 = 5; $time2<=$prnum; $lt++,$time2++){
        $lad = $lan + $maxnum + 1;
        //$col = chr(97 + $x);
        $cola = chr(66 + ($lt * 2)); //66 =B;
        $colb = chr(67 + ($lt * 2)); //66 =B;
        //echo '第一行'.$col.$i;
        $spreadsheet->getActiveSheet()->getColumnDimension($cola)->setWidth(20);  //列宽度
        $spreadsheet->getActiveSheet()->getColumnDimension($colb)->setWidth(20);  //列宽度

        $spreadsheet->getActiveSheet()->getStyle("{$cola}1:{$colb}1")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue($cola.'1', $cpsall['cpsp2'][$lan][0]["cpsno"]);

        $spreadsheet->getActiveSheet()->getStyle("{$colb}2")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue("{$colb}2", $cpsall['cpsp2'][$lad][4]);
        /*
    //$spreadsheet->getActiveSheet()->setCellValue('C1', $cpsall['cpsp2']["maxnum"]);
        $spreadsheet->getActiveSheet()->mergeCells("{$cola}3:{$colb}3");
        $spreadsheet->getActiveSheet()->getStyle("{$cola}3:{$colb}3")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}3", $cpsall['cpsp2'][$lan][0]["ftyno"]);
    */
        $spreadsheet->getActiveSheet()->getStyle("{$cola}3:{$colb}3")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->mergeCells("{$cola}3:{$colb}3");
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}3", $cpsall['cpsp2'][$lan][0]["jobno"]);

        $spreadsheet->getActiveSheet()->getStyle("{$cola}4:{$colb}4")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->mergeCells("{$cola}4:{$colb}4");
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}4", $cpsall['cpsp2'][$lan][0]["styleno"]);



        for($listtop = 0,$lttime1 = 0,$listta = 5;$lttime1<=9;$listtop++,$listta++,$lttime1++){
            if($lttime1 == 4)
            {$listta = 8;
                continue;};
            $spreadsheet->getActiveSheet()->getStyle("{$cola}{$listta}:{$colb}{$listta}")->applyFromArray($styleArray1);
            $spreadsheet->getActiveSheet()->mergeCells("{$cola}{$listta}:{$colb}{$listta}");
            $spreadsheet->getActiveSheet()->setCellValue("{$cola}{$listta}", $cpsall['cpsp2'][$lad][$listtop]);
        }


        /*加載圖片*/
        $img = $cpsall['cpsp2'][$lan][0]["remarkimg2"];
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


// Generate an image
//$gdImage = @imagecreatetruecolor($width, $height) or die('Cannot Initialize new GD image stream');
//$textColor = imagecolorallocate($gdImage, 255, 255, 255);
//imagestring($gdImage, 1, 5, 5,  'Created with PhpSpreadsheet', $textColor);

// Add a drawing to the worksheet
        $drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
        $drawing->setName($cpsall['cpsp2'][$lad][4]);
        $drawing->setDescription($cpsall['cpsp2'][$lad][4]);
//$drawing->setImageResource($gdImage);
        $drawing->setImageResource($img);
        $drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
        $drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
//$drawing->setHeight($width);

        $drawing->setWidth($width>120 ? 120:$width);
//$drawing->setWidth(180);
//$drawing->setHeight(150);
        $drawing->setCoordinates($cola.'2');
        $drawing->setOffsetX(5);
        $drawing->setOffsetY(5);
        $drawing->setWorksheet($spreadsheet->getActiveSheet());

        /*加載圖片*/

        for($j = 0 ,$z=10, $x = 14; $j < 13 ; $j++) {
            //$z=//數組值
            //$x = 20; //行數

            for ($i = 0; $i < 2; $i++) {
                $list = chr(66 + $i + ($lt * 2)); //66 =B;
                $spreadsheet->getActiveSheet()->getRowDimension($x)->setRowHeight(32); //列高度
                $spreadsheet->getActiveSheet()->getStyle($list.$x)->applyFromArray($styleArray1);
                $spreadsheet->getActiveSheet()->setCellValue($list.$x,$cpsall['cpsp2'][$lad][$z] );
                $z++;
            }
            $x++;
        }

        $lan++;

    } //1st for

}
$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

/**
 * 第三頁
 */
$prnum = $maxnum <=14 ? $maxnum : 14;
if($maxnum > 9 ){

    /**
     * $lt 列名
     * $lan 列數組序號 取數據
     * $lad 列數組數據 $lan + $prnum+1
     */
    $opensheetnum++;
    $spreadsheet->setActiveSheetIndex($opensheetnum);  //設置當前活動表

    /*A列名称*/
    $fabname=array("CPS","Sketch：","Job no.",'Style no','Shipment date','Net Weight(g)','Fabric composition','Lining','Trim fabric','物料','訂布用料','最新用料（Y/件）','特殊工序','Lab dips submitted','Lab dips approved','Bulk cloth approved','Bulk test report approved','Bulk fabric ready in factory','上布方式','Care label','Blue seal approved','Green Seal approved','White seal lastest submit date','Remarks','QC & M.Detector report','SHORT SHIPPED');
    $fabcou = count($fabname);
    for ($j=0,$row = 1;$j<$fabcou;$j++,$row++) {
        $spreadsheet->getActiveSheet()->setCellValue('A'.$row, $fabname[$j]);
        $spreadsheet->getActiveSheet()->getStyle('A'.$row)->applyFromArray($styleArray1);
    }

    $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(20);  //列宽度
    /*A列名称*/

    for($lt = 0, $lan = 10,$time3 = 10; $time3<=$prnum; $lt++,$time3++){
        $lad = $lan + $maxnum + 1;
        //$col = chr(97 + $x);
        $cola = chr(66 + ($lt * 2)); //66 =B;
        $colb = chr(67 + ($lt * 2)); //66 =B;
        //echo '第一行'.$col.$i;
        $spreadsheet->getActiveSheet()->getColumnDimension($cola)->setWidth(20);  //列宽度
        $spreadsheet->getActiveSheet()->getColumnDimension($colb)->setWidth(20);  //列宽度

        $spreadsheet->getActiveSheet()->getStyle("{$cola}1:{$colb}1")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue($cola.'1', $cpsall['cpsp2'][$lan][0]["cpsno"]);

        $spreadsheet->getActiveSheet()->getStyle("{$colb}2")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue("{$colb}2", $cpsall['cpsp2'][$lad][4]);
        /*
    //$spreadsheet->getActiveSheet()->setCellValue('C1', $cpsall['cpsp2']["maxnum"]);
        $spreadsheet->getActiveSheet()->mergeCells("{$cola}3:{$colb}3");
        $spreadsheet->getActiveSheet()->getStyle("{$cola}3:{$colb}3")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}3", $cpsall['cpsp2'][$lan][0]["ftyno"]);
    */
        $spreadsheet->getActiveSheet()->getStyle("{$cola}3:{$colb}3")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->mergeCells("{$cola}3:{$colb}3");
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}3", $cpsall['cpsp2'][$lan][0]["jobno"]);

        $spreadsheet->getActiveSheet()->getStyle("{$cola}4:{$colb}4")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->mergeCells("{$cola}4:{$colb}4");
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}4", $cpsall['cpsp2'][$lan][0]["styleno"]);



        for($listtop = 0,$lttime1 = 0,$listta = 5;$lttime1<=9;$listtop++,$listta++,$lttime1++){
            if($lttime1 == 4)
            {$listta = 8;
                continue;};
            $spreadsheet->getActiveSheet()->getStyle("{$cola}{$listta}:{$colb}{$listta}")->applyFromArray($styleArray1);
            $spreadsheet->getActiveSheet()->mergeCells("{$cola}{$listta}:{$colb}{$listta}");
            $spreadsheet->getActiveSheet()->setCellValue("{$cola}{$listta}", $cpsall['cpsp2'][$lad][$listtop]);
        }


        /*加載圖片*/
        $img = $cpsall['cpsp2'][$lan][0]["remarkimg2"];
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


// Generate an image
//$gdImage = @imagecreatetruecolor($width, $height) or die('Cannot Initialize new GD image stream');
//$textColor = imagecolorallocate($gdImage, 255, 255, 255);
//imagestring($gdImage, 1, 5, 5,  'Created with PhpSpreadsheet', $textColor);

// Add a drawing to the worksheet
        $drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
        $drawing->setName($cpsall['cpsp2'][$lad][4]);
        $drawing->setDescription($cpsall['cpsp2'][$lad][4]);
//$drawing->setImageResource($gdImage);
        $drawing->setImageResource($img);
        $drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
        $drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
//$drawing->setHeight($width);

        $drawing->setWidth($width>120 ? 120:$width);
//$drawing->setWidth(180);
//$drawing->setHeight(150);
        $drawing->setCoordinates($cola.'2');
        $drawing->setOffsetX(5);
        $drawing->setOffsetY(5);
        $drawing->setWorksheet($spreadsheet->getActiveSheet());

        /*加載圖片*/

        for($j = 0 ,$z=10, $x = 14; $j < 13 ; $j++) {
            //$z=//數組值
            //$x = 20; //行數

            for ($i = 0; $i < 2; $i++) {
                $list = chr(66 + $i + ($lt * 2)); //66 =B;
                $spreadsheet->getActiveSheet()->getRowDimension($x)->setRowHeight(32); //列高度
                $spreadsheet->getActiveSheet()->getStyle($list.$x)->applyFromArray($styleArray1);
                $spreadsheet->getActiveSheet()->setCellValue($list.$x,$cpsall['cpsp2'][$lad][$z] );
                $z++;
            }
            $x++;
        }

        $lan++;

    } //1st for

}

$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页
/*CPSP2  CPS第二份*/

/**
 * CPSP3  CPS第三份
 */
$opensheetnum++;
$spreadsheet->setActiveSheetIndex($opensheetnum);  //設置當前活動表


$sheet = $spreadsheet->getActiveSheet();



$spreadsheet->getDefaultStyle()->getFont()->setName('微软雅黑');
$spreadsheet->getDefaultStyle()->getFont()->setSize(10);

/*A列名称*/
$fabname=array("CPS","Sketch："," ",'Style no','Fabrication','Fabric Price','Cons./Doz (NET)','Lining','Lining Price','Cons./Doz (NET)','Contrast fabric','Contrast Price','Cons./Doz (NET)','FABRIC COST','interlinging @8/y','THERAD','Zipper','Eyes&Hooks','Button','Elastic','Grosgrain','Pintuck/打条','Hanger loop','Label','Hangtag','Polybag','Hanger','Carton','Total Trim Cost','Sewing(車縫工價)','Cut,Trim,Pack etc.','QTY','FOB USD');
$fabcou = count($fabname);
for ($j=0,$row = 1;$j<$fabcou;$j++,$row++) {
    $spreadsheet->getActiveSheet()->setCellValue('A'.$row, $fabname[$j]);
    $spreadsheet->getActiveSheet()->getStyle('A'.$row)->applyFromArray($styleArray1);
}

$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(20);  //列宽度
/*A列名称*/


$styleArray1 = [
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
        'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP,
        'wrapText' => true,
        'ShrinkToFit'=>true,
    ],
    'font' => [
        'Size' => '10',
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

$maxnum = $cpsall['cpsp3']["maxnum"];


$prnum = $maxnum < 5 ? $maxnum : 4;

/**
 * $lt 列名
 * $lan 列數組序號
 * $lad 列數組數據 $lan + $prnum+1
 */


$spreadsheet->getActiveSheet()->getRowDimension('2')->setRowHeight(100); //列高度


for($lt = 0, $lan = 0; $lt<=$prnum; $lt++){
    $lad = $lan + $maxnum + 1;
    //$col = chr(97 + $x);
    $cola = chr(66 + ($lt * 1)); //66 =B;
    //$colb = chr(67 + ($lt * 2)); //66 =B;
    //echo '第一行'.$col.$i;
    $spreadsheet->getActiveSheet()->getColumnDimension($cola)->setWidth(20);  //列宽度


    $spreadsheet->getActiveSheet()->getStyle("{$cola}1")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->setCellValue($cola.'1', $cpsall['cpsp3'][$lan][0]["cpsno"]);

    $spreadsheet->getActiveSheet()->getStyle("{$cola}3")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->setCellValue("{$cola}3", $cpsall['cpsp3'][$lad][0]);
    /*
//$spreadsheet->getActiveSheet()->setCellValue('C1', $cpsall['cpsp3']["maxnum"]);
    $spreadsheet->getActiveSheet()->mergeCells("{$cola}3:{$colb}3");
    $spreadsheet->getActiveSheet()->getStyle("{$cola}3:{$colb}3")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->setCellValue("{$cola}3", $cpsall['cpsp3'][$lan][0]["ftyno"]);
*/


    $spreadsheet->getActiveSheet()->getStyle("{$cola}4")->applyFromArray($styleArray1);

    $spreadsheet->getActiveSheet()->setCellValue("{$cola}4", $cpsall['cpsp3'][$lan][0]["styleno"]);



    for($listtop = 1,$lttime1 = 1,$listta = 5;$lttime1<=29;$listtop++,$listta++,$lttime1++){
        $spreadsheet->getActiveSheet()->getRowDimension($listta)->setRowHeight(32); //列高度
        $spreadsheet->getActiveSheet()->getStyle("{$cola}{$listta}")->applyFromArray($styleArray1);

        $spreadsheet->getActiveSheet()->setCellValue("{$cola}{$listta}", $cpsall['cpsp3'][$lad][$listtop]);
    }


    /*加載圖片*/
    $img = $cpsall['cpsp3'][$lan][0]["remarkimg2"];
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


// Generate an image
//$gdImage = @imagecreatetruecolor($width, $height) or die('Cannot Initialize new GD image stream');
//$textColor = imagecolorallocate($gdImage, 255, 255, 255);
//imagestring($gdImage, 1, 5, 5,  'Created with PhpSpreadsheet', $textColor);

// Add a drawing to the worksheet
    $drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
    $drawing->setName($cpsall['cpsp3'][$lad][4]);
    $drawing->setDescription($cpsall['cpsp3'][$lad][4]);
//$drawing->setImageResource($gdImage);
    $drawing->setImageResource($img);
    $drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
    $drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
//$drawing->setHeight($width);

    $drawing->setHeight($height>120 ? 120:$height);
//$drawing->setWidth(180);
//$drawing->setHeight(150);
    $drawing->setCoordinates($cola.'2');
    $drawing->setOffsetX(5);
    $drawing->setOffsetY(5);
    $drawing->setWorksheet($spreadsheet->getActiveSheet());

    /*加載圖片*/


    $lan++;
} //1st for
$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

/**
 *  cps3 第二頁
 */
$prnum = $maxnum <= 9 ? $maxnum : 9;
if($maxnum > 4 ){

    /**
     * $lt 列名
     * $lan 列數組序號
     * $lad 列數組數據 $lan + $prnum+1
     */
    $opensheetnum++;
    $spreadsheet->setActiveSheetIndex($opensheetnum);  //設置當前活動表


    $spreadsheet->getActiveSheet()->getRowDimension('2')->setRowHeight(100); //列高度

    /*A列名称*/
    $fabname=array("CPS","Sketch："," ",'Style no','Fabrication','Fabric Price','Cons./Doz (NET)','Lining','Lining Price','Cons./Doz (NET)','Contrast fabric','Contrast Price','Cons./Doz (NET)','FABRIC COST','interlinging @8/y','THERAD','Zipper','Eyes&Hooks','Button','Elastic','Grosgrain','Pintuck/打条','Hanger loop','Label','Hangtag','Polybag','Hanger','Carton','Total Trim Cost','Sewing(車縫工價)','Cut,Trim,Pack etc.','QTY','FOB USD');
    $fabcou = count($fabname);
    for ($j=0,$row = 1;$j<$fabcou;$j++,$row++) {
        $spreadsheet->getActiveSheet()->setCellValue('A'.$row, $fabname[$j]);
        $spreadsheet->getActiveSheet()->getStyle('A'.$row)->applyFromArray($styleArray1);
    }

    $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(20);  //列宽度
    /*A列名称*/


    for($lt = 0, $lan = 0; $lt<=$prnum; $lt++){
        $lad = $lan + $maxnum + 1;
        //$col = chr(97 + $x);
        $cola = chr(66 + ($lt * 1)); //66 =B;
        //$colb = chr(67 + ($lt * 2)); //66 =B;
        //echo '第一行'.$col.$i;
        $spreadsheet->getActiveSheet()->getColumnDimension($cola)->setWidth(20);  //列宽度


        $spreadsheet->getActiveSheet()->getStyle("{$cola}1")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue($cola.'1', $cpsall['cpsp3'][$lan][0]["cpsno"]);

        $spreadsheet->getActiveSheet()->getStyle("{$cola}3")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}3", $cpsall['cpsp3'][$lad][0]);
        /*
    //$spreadsheet->getActiveSheet()->setCellValue('C1', $cpsall['cpsp3']["maxnum"]);
        $spreadsheet->getActiveSheet()->mergeCells("{$cola}3:{$colb}3");
        $spreadsheet->getActiveSheet()->getStyle("{$cola}3:{$colb}3")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}3", $cpsall['cpsp3'][$lan][0]["ftyno"]);
    */


        $spreadsheet->getActiveSheet()->getStyle("{$cola}4")->applyFromArray($styleArray1);

        $spreadsheet->getActiveSheet()->setCellValue("{$cola}4", $cpsall['cpsp3'][$lan][0]["styleno"]);



        for($listtop = 1,$lttime1 = 1,$listta = 5;$lttime1<=29;$listtop++,$listta++,$lttime1++){
            $spreadsheet->getActiveSheet()->getRowDimension($listta)->setRowHeight(32); //列高度
            $spreadsheet->getActiveSheet()->getStyle("{$cola}{$listta}")->applyFromArray($styleArray1);

            $spreadsheet->getActiveSheet()->setCellValue("{$cola}{$listta}", $cpsall['cpsp3'][$lad][$listtop]);
        }


        /*加載圖片*/
        $img = $cpsall['cpsp3'][$lan][0]["remarkimg2"];
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


// Generate an image
//$gdImage = @imagecreatetruecolor($width, $height) or die('Cannot Initialize new GD image stream');
//$textColor = imagecolorallocate($gdImage, 255, 255, 255);
//imagestring($gdImage, 1, 5, 5,  'Created with PhpSpreadsheet', $textColor);

// Add a drawing to the worksheet
        $drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
        $drawing->setName($cpsall['cpsp3'][$lad][4]);
        $drawing->setDescription($cpsall['cpsp3'][$lad][4]);
//$drawing->setImageResource($gdImage);
        $drawing->setImageResource($img);
        $drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
        $drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
//$drawing->setHeight($width);

        $drawing->setHeight($height>120 ? 120:$height);
//$drawing->setWidth(180);
//$drawing->setHeight(150);
        $drawing->setCoordinates($cola.'2');
        $drawing->setOffsetX(5);
        $drawing->setOffsetY(5);
        $drawing->setWorksheet($spreadsheet->getActiveSheet());

        /*加載圖片*/


        $lan++;

    } //1st for

}
$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页


/**
 * 第三頁
 */
$prnum = $maxnum <=14 ? $maxnum : 14;
if($maxnum > 9 ){

    /**
     * $lt 列名
     * $lan 列數組序號
     * $lad 列數組數據 $lan + $prnum+1
     */
    $opensheetnum++;
    $spreadsheet->setActiveSheetIndex($opensheetnum);  //設置當前活動表

    $spreadsheet->getActiveSheet()->getRowDimension('2')->setRowHeight(100); //列高度

    /*A列名称*/
    $fabname=array("CPS","Sketch："," ",'Style no','Fabrication','Fabric Price','Cons./Doz (NET)','Lining','Lining Price','Cons./Doz (NET)','Contrast fabric','Contrast Price','Cons./Doz (NET)','FABRIC COST','interlinging @8/y','THERAD','Zipper','Eyes&Hooks','Button','Elastic','Grosgrain','Pintuck/打条','Hanger loop','Label','Hangtag','Polybag','Hanger','Carton','Total Trim Cost','Sewing(車縫工價)','Cut,Trim,Pack etc.','QTY','FOB USD');
    $fabcou = count($fabname);
    for ($j=0,$row = 1;$j<$fabcou;$j++,$row++) {
        $spreadsheet->getActiveSheet()->setCellValue('A'.$row, $fabname[$j]);
        $spreadsheet->getActiveSheet()->getStyle('A'.$row)->applyFromArray($styleArray1);
    }

    $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(20);  //列宽度
    /*A列名称*/


    for($lt = 0, $lan = 0; $lt<=$prnum; $lt++){
        $lad = $lan + $maxnum + 1;
        //$col = chr(97 + $x);
        $cola = chr(66 + ($lt * 1)); //66 =B;
        //$colb = chr(67 + ($lt * 2)); //66 =B;
        //echo '第一行'.$col.$i;
        $spreadsheet->getActiveSheet()->getColumnDimension($cola)->setWidth(20);  //列宽度


        $spreadsheet->getActiveSheet()->getStyle("{$cola}1")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue($cola.'1', $cpsall['cpsp3'][$lan][0]["cpsno"]);

        $spreadsheet->getActiveSheet()->getStyle("{$cola}3")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}3", $cpsall['cpsp3'][$lad][0]);
        /*
    //$spreadsheet->getActiveSheet()->setCellValue('C1', $cpsall['cpsp3']["maxnum"]);
        $spreadsheet->getActiveSheet()->mergeCells("{$cola}3:{$colb}3");
        $spreadsheet->getActiveSheet()->getStyle("{$cola}3:{$colb}3")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}3", $cpsall['cpsp3'][$lan][0]["ftyno"]);
    */


        $spreadsheet->getActiveSheet()->getStyle("{$cola}4")->applyFromArray($styleArray1);

        $spreadsheet->getActiveSheet()->setCellValue("{$cola}4", $cpsall['cpsp3'][$lan][0]["styleno"]);



        for($listtop = 1,$lttime1 = 1,$listta = 5;$lttime1<=29;$listtop++,$listta++,$lttime1++){
            $spreadsheet->getActiveSheet()->getRowDimension($listta)->setRowHeight(32); //列高度
            $spreadsheet->getActiveSheet()->getStyle("{$cola}{$listta}")->applyFromArray($styleArray1);

            $spreadsheet->getActiveSheet()->setCellValue("{$cola}{$listta}", $cpsall['cpsp3'][$lad][$listtop]);
        }


        /*加載圖片*/
        $img = $cpsall['cpsp3'][$lan][0]["remarkimg2"];
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


// Generate an image
//$gdImage = @imagecreatetruecolor($width, $height) or die('Cannot Initialize new GD image stream');
//$textColor = imagecolorallocate($gdImage, 255, 255, 255);
//imagestring($gdImage, 1, 5, 5,  'Created with PhpSpreadsheet', $textColor);

// Add a drawing to the worksheet
        $drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
        $drawing->setName($cpsall['cpsp3'][$lad][4]);
        $drawing->setDescription($cpsall['cpsp3'][$lad][4]);
//$drawing->setImageResource($gdImage);
        $drawing->setImageResource($img);
        $drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
        $drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
//$drawing->setHeight($width);

        $drawing->setHeight($height>110 ? 110:$height);
//$drawing->setWidth(180);
//$drawing->setHeight(150);
        $drawing->setCoordinates($cola.'2');
        $drawing->setOffsetX(5);
        $drawing->setOffsetY(5);
        $drawing->setWorksheet($spreadsheet->getActiveSheet());

        /*加載圖片*/


        $lan++;

    } //1st for

}

$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

/*CPSP3 第三份*/



unset($_SESSION['cpsall'] ); //注销SESSION
// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$spreadsheet->setActiveSheetIndex(0);

$output=  ($_GET['action'] == 'formdown' )? 1:0;
$nt = date("YmdHis",time()); //转换为日期。
$filenameout = 'cpsallout'.$nt.'.xlsx';
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
