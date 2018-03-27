<?php
session_start();
global $action;

$cpsp2 =  $_SESSION['cpsp2'];
//var_dump($cpsp2);
//require '/home/soft/vendor/autoload.php';
//require '../vendor/autoload.php';
require '/home/pan/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;


    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/cpsp2.xlsx');

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

$maxnum = $cpsp2["maxnum"];


$prnum = $maxnum < 5 ? $maxnum : 4;

/**
 * $lt 列名
 * $lan 列數組序號
 * $lad 列數組數據 $lan + $prnum+1
 */
$spreadsheet->setActiveSheetIndex(0);  //設置當前活動表

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
    $spreadsheet->getActiveSheet()->setCellValue($cola.'1', $cpsp2[$lan][0]["cpsno"]);

    $spreadsheet->getActiveSheet()->getStyle("{$colb}2")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->setCellValue("{$colb}2", $cpsp2[$lad][4]);
    /*
//$spreadsheet->getActiveSheet()->setCellValue('C1', $cpsp2["maxnum"]);
    $spreadsheet->getActiveSheet()->mergeCells("{$cola}3:{$colb}3");
    $spreadsheet->getActiveSheet()->getStyle("{$cola}3:{$colb}3")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->setCellValue("{$cola}3", $cpsp2[$lan][0]["ftyno"]);
*/
    $spreadsheet->getActiveSheet()->getStyle("{$cola}3:{$colb}3")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->mergeCells("{$cola}3:{$colb}3");
    $spreadsheet->getActiveSheet()->setCellValue("{$cola}3", $cpsp2[$lan][0]["jobno"]);

    $spreadsheet->getActiveSheet()->getStyle("{$cola}4:{$colb}4")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->mergeCells("{$cola}4:{$colb}4");
    $spreadsheet->getActiveSheet()->setCellValue("{$cola}4", $cpsp2[$lan][0]["styleno"]);



    for($listtop = 0,$lttime1 = 0,$listta = 5;$lttime1<=9;$listtop++,$listta++,$lttime1++){
            if($lttime1 == 4)
            {$listta = 8;
                continue;};
        $spreadsheet->getActiveSheet()->getStyle("{$cola}{$listta}:{$colb}{$listta}")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->mergeCells("{$cola}{$listta}:{$colb}{$listta}");
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}{$listta}", $cpsp2[$lad][$listtop]);
    }


    /*加載圖片*/
    $img = $cpsp2[$lan][0]["remarkimg2"];
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
    $drawing->setName($cpsp2[$lad][4]);
    $drawing->setDescription($cpsp2[$lad][4]);
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
            $spreadsheet->getActiveSheet()->setCellValue($list.$x,$cpsp2[$lad][$z] );
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

    /**
     * $lt 列名
     * $lan 列數組序號 取數據
     * $lad 列數組數據 $lan + $prnum+1
     */
    $spreadsheet->setActiveSheetIndex(1);  //設置當前活動表


    for($lt = 0, $lan = 5,$time2 = 5; $time2<=$prnum; $lt++,$time2++){
        $lad = $lan + $maxnum + 1;
        //$col = chr(97 + $x);
        $cola = chr(66 + ($lt * 2)); //66 =B;
        $colb = chr(67 + ($lt * 2)); //66 =B;
        //echo '第一行'.$col.$i;
        $spreadsheet->getActiveSheet()->getColumnDimension($cola)->setWidth(20);  //列宽度
        $spreadsheet->getActiveSheet()->getColumnDimension($colb)->setWidth(20);  //列宽度

        $spreadsheet->getActiveSheet()->getStyle("{$cola}1:{$colb}1")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue($cola.'1', $cpsp2[$lan][0]["cpsno"]);

        $spreadsheet->getActiveSheet()->getStyle("{$colb}2")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue("{$colb}2", $cpsp2[$lad][4]);
        /*
    //$spreadsheet->getActiveSheet()->setCellValue('C1', $cpsp2["maxnum"]);
        $spreadsheet->getActiveSheet()->mergeCells("{$cola}3:{$colb}3");
        $spreadsheet->getActiveSheet()->getStyle("{$cola}3:{$colb}3")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}3", $cpsp2[$lan][0]["ftyno"]);
    */
        $spreadsheet->getActiveSheet()->getStyle("{$cola}3:{$colb}3")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->mergeCells("{$cola}3:{$colb}3");
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}3", $cpsp2[$lan][0]["jobno"]);

        $spreadsheet->getActiveSheet()->getStyle("{$cola}4:{$colb}4")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->mergeCells("{$cola}4:{$colb}4");
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}4", $cpsp2[$lan][0]["styleno"]);



        for($listtop = 0,$lttime1 = 0,$listta = 5;$lttime1<=9;$listtop++,$listta++,$lttime1++){
            if($lttime1 == 4)
            {$listta = 8;
                continue;};
            $spreadsheet->getActiveSheet()->getStyle("{$cola}{$listta}:{$colb}{$listta}")->applyFromArray($styleArray1);
            $spreadsheet->getActiveSheet()->mergeCells("{$cola}{$listta}:{$colb}{$listta}");
            $spreadsheet->getActiveSheet()->setCellValue("{$cola}{$listta}", $cpsp2[$lad][$listtop]);
        }


        /*加載圖片*/
        $img = $cpsp2[$lan][0]["remarkimg2"];
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
        $drawing->setName($cpsp2[$lad][4]);
        $drawing->setDescription($cpsp2[$lad][4]);
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
                $spreadsheet->getActiveSheet()->setCellValue($list.$x,$cpsp2[$lad][$z] );
                $z++;
            }
            $x++;
        }

        $lan++;

    } //1st for

}


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
    $spreadsheet->setActiveSheetIndex(2);  //設置當前活動表

    for($lt = 0, $lan = 10,$time3 = 10; $time3<=$prnum; $lt++,$time3++){
        $lad = $lan + $maxnum + 1;
        //$col = chr(97 + $x);
        $cola = chr(66 + ($lt * 2)); //66 =B;
        $colb = chr(67 + ($lt * 2)); //66 =B;
        //echo '第一行'.$col.$i;
        $spreadsheet->getActiveSheet()->getColumnDimension($cola)->setWidth(20);  //列宽度
        $spreadsheet->getActiveSheet()->getColumnDimension($colb)->setWidth(20);  //列宽度

        $spreadsheet->getActiveSheet()->getStyle("{$cola}1:{$colb}1")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue($cola.'1', $cpsp2[$lan][0]["cpsno"]);

        $spreadsheet->getActiveSheet()->getStyle("{$colb}2")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue("{$colb}2", $cpsp2[$lad][4]);
        /*
    //$spreadsheet->getActiveSheet()->setCellValue('C1', $cpsp2["maxnum"]);
        $spreadsheet->getActiveSheet()->mergeCells("{$cola}3:{$colb}3");
        $spreadsheet->getActiveSheet()->getStyle("{$cola}3:{$colb}3")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}3", $cpsp2[$lan][0]["ftyno"]);
    */
        $spreadsheet->getActiveSheet()->getStyle("{$cola}3:{$colb}3")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->mergeCells("{$cola}3:{$colb}3");
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}3", $cpsp2[$lan][0]["jobno"]);

        $spreadsheet->getActiveSheet()->getStyle("{$cola}4:{$colb}4")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->mergeCells("{$cola}4:{$colb}4");
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}4", $cpsp2[$lan][0]["styleno"]);



        for($listtop = 0,$lttime1 = 0,$listta = 5;$lttime1<=9;$listtop++,$listta++,$lttime1++){
            if($lttime1 == 4)
            {$listta = 8;
                continue;};
            $spreadsheet->getActiveSheet()->getStyle("{$cola}{$listta}:{$colb}{$listta}")->applyFromArray($styleArray1);
            $spreadsheet->getActiveSheet()->mergeCells("{$cola}{$listta}:{$colb}{$listta}");
            $spreadsheet->getActiveSheet()->setCellValue("{$cola}{$listta}", $cpsp2[$lad][$listtop]);
        }


        /*加載圖片*/
        $img = $cpsp2[$lan][0]["remarkimg2"];
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
        $drawing->setName($cpsp2[$lad][4]);
        $drawing->setDescription($cpsp2[$lad][4]);
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
                $spreadsheet->getActiveSheet()->setCellValue($list.$x,$cpsp2[$lad][$z] );
                $z++;
            }
            $x++;
        }

        $lan++;

    } //1st for

}

$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页


unset($_SESSION['cpsp2'] ); //注销SESSION

$output=  ($_GET['action'] == 'formdown' )? 1:0;
$nt = date("YmdHis",time()); //转换为日期。
$filenameout = 'cpsp2out'.$nt.'.xlsx';
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
