<?php
session_start();

$cpsp2 =  $_SESSION['cpsall'];

//var_dump($cpsp2);

//require '../vendor/autoload.php';
//require '/home/pan/vendor/autoload.php';

require_once('autoloadconfig.php');  //判断是否在线

if($online){
    require_once '/home/pan/vendor/autoload.php';

}else{
    require_once '/Applications/XAMPP/xamppfiles/htdocs/composer/vendor/autoload.php';
}
require_once ('img.php');
use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;


$spreadsheet = new Spreadsheet();
//$sheet = $spreadsheet->getActiveSheet(0);

$spreadsheet->getDefaultStyle()->getFont()->setName('微软雅黑');
$spreadsheet->getDefaultStyle()->getFont()->setSize(12);

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

$maxnum = $cpsp2["arrcount"] - 1;
$maxlist = $cpsp2["arrcount"] ;


$stotal = (ceil($maxlist / 5) > 1) ? (ceil($maxlist / 5)  ) : 1;  //总共有多少页？


for($spage = 0;$spage< $stotal;$spage++){
    if($spage>0){
        $worksheet1 = $spreadsheet->createSheet(); //新增sheet；
        //$spreadsheet->setActiveSheetIndex($spage);  //設置當前活動表
    }
    $spreadsheet->setActiveSheetIndex($spage);  //設置當前活動表
    $sname = $spage +1 ;
    $spreadsheet->getActiveSheet()->setTitle("cpsp2 sheet".$sname);


    if($spage > 0){
        $prnum = $maxnum <= (4 + 5 * $spage) ? $maxnum : (4 + 5 * $spage);
        //$prnum = $maxnum <= 9 ? $maxnum : 9;
    }else{
        $prnum = $maxnum < 5 ? $maxnum : 4;
    }
//$prnum = $maxnum < (5 + 5 * $spage) ? $maxnum : (4 + 5 * $spage);
//$prnum = $maxnum <= 9 ? $maxnum : 9;
    /**
     * $lt 列名
     * $lan 列數組序號
     * $lad 列數組數據 $lan + $prnum+1
     */

$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(16);  //列宽度
$spreadsheet->getActiveSheet()->setCellValue('A1', 'cpsno');

$spreadsheet->getActiveSheet()->setCellValue('A2', 'sketch');
$spreadsheet->getActiveSheet()->setCellValue('A3', 'Fty no.：');
$spreadsheet->getActiveSheet()->setCellValue('A4', 'Job no.：');
$spreadsheet->getActiveSheet()->setCellValue('A5', 'Style no：');
$spreadsheet->getActiveSheet()->setCellValue('A6', 'Shipment date：');
$spreadsheet->getActiveSheet()->setCellValue('A7', 'Net Weight(g)：');
$spreadsheet->getActiveSheet()->setCellValue('A8', 'Fabric composition：');
$spreadsheet->getActiveSheet()->setCellValue('A9', 'Lining：');
$spreadsheet->getActiveSheet()->setCellValue('A10', 'Trim fabric：');
$spreadsheet->getActiveSheet()->setCellValue('A11', '物料：');
$spreadsheet->getActiveSheet()->setCellValue('A12', '訂布用料：');
$spreadsheet->getActiveSheet()->setCellValue('A13', '最新用料（Y/件）：');
$spreadsheet->getActiveSheet()->setCellValue('A14', '特殊工序：');


    $spreadsheet->getActiveSheet()->getRowDimension('2')->setRowHeight(100); //列高度
    for ($i = 6 ;$i <= 14 ; $i++){
        $spreadsheet->getActiveSheet()->getRowDimension($i)->setRowHeight(32); //列高度
    }



    if($spage == 0){
        $nowprnum =  $prnum ;
        $lan = 0;
    }else{
        $nowprnum =    $prnum  - (5 * $spage) ;
        $lan = (5 * $spage);
    }

    $maxitem = 0;    //表格中 item 行数最大的数值是
    for($q = 0,$n = $lan; $q<= $nowprnum ;$q++,$n++){
        $itemarr = array();
        $itemarr = json_decode(stripcslashes($cpsp2[$n][0]["item"]), true);
        $itemcol = (count($itemarr)/3);
        $maxitem =  ($maxitem >= $itemcol ? $maxitem : $itemcol ) ;
    }



    for($lt = 0; $lt <= $nowprnum ; $lt++,$lan++) {
        $lad = $lan + $maxnum + 1;
        //$col = chr(97 + $x);
        $cola = chr(66 + ($lt * 3)); //66 =B;
        $colb = chr(67 + ($lt * 3)); //67 =C;
        $colc = chr(68 + ($lt * 3)); //68 =D;
        //echo '第一行'.$col.$i;
        $spreadsheet->getActiveSheet()->getColumnDimension($cola)->setWidth(16);  //列宽度
        $spreadsheet->getActiveSheet()->getColumnDimension($colb)->setWidth(16);  //列宽度
        $spreadsheet->getActiveSheet()->getColumnDimension($colc)->setWidth(16);  //列宽度

        $spreadsheet->getActiveSheet()->getStyle("{$cola}1:{$colc}1")->applyFromArray($styleArray1);

        $spreadsheet->getActiveSheet()->setCellValue($cola . '1', $cpsp2[$lan][0]["cpsno"]);
//        $spreadsheet->getActiveSheet()->setCellValue($cola . '1', $n);
//        $spreadsheet->getActiveSheet()->setCellValue($colb . '1', $nowprnum);
//        $spreadsheet->getActiveSheet()->setCellValue($colc . '1', $lt);

        $spreadsheet->getActiveSheet()->getStyle("{$cola}3:{$colc}3")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->mergeCells("{$cola}3:{$colc}3");
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}3", $cpsp2[$lan][0]["ftyno"]);
        $spreadsheet->getActiveSheet()->getStyle("{$cola}3:{$colc}3")->getAlignment()->setShrinkToFit(true);//缩小以适合

        $spreadsheet->getActiveSheet()->getStyle("{$cola}4:{$colc}4")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->mergeCells("{$cola}4:{$colc}4");
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}4", $cpsp2[$lan][0]["jobno"]);
        $spreadsheet->getActiveSheet()->getStyle("{$cola}4:{$colc}4")->getAlignment()->setShrinkToFit(true);//缩小以适合

        $spreadsheet->getActiveSheet()->getStyle("{$cola}5:{$colc}5")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->mergeCells("{$cola}5:{$colc}5");
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}5", $cpsp2[$lan][0]["styleno"]);
        $spreadsheet->getActiveSheet()->getStyle("{$cola}5:{$colc}5")->getAlignment()->setShrinkToFit(true);//缩小以适合


        $fabarr = array();
        $fabarr = json_decode(stripcslashes($cpsp2[$lan][0]["fab"]), true);

        $itemarr = array();
        $itemarr = json_decode(stripcslashes($cpsp2[$lan][0]["item"]), true);

    for($t = 1,$l = 6;$t<=9;$t++,$l++){

        $spreadsheet->getActiveSheet()->getStyle("{$cola}{$l}:{$colc}{$l}")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->mergeCells("{$cola}{$l}:{$colc}{$l}");
        $spreadsheet->getActiveSheet()->setCellValue("{$cola}{$l}", $fabarr[$t]);
        $spreadsheet->getActiveSheet()->getStyle("{$cola}{$l}:{$colc}{$l}")->getAlignment()->setShrinkToFit(true);//缩小以适合
    }




/* 图片标注*/
    $spreadsheet->getActiveSheet()->getStyle("{$colb}2:{$colc}2")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->mergeCells("{$colb}2:{$colc}2");
    $spreadsheet->getActiveSheet()->setCellValue($colb.'2', $fabarr[0]);
    $spreadsheet->getActiveSheet()->getStyle("{$colb}2:{$colc}2")->getAlignment()->setShrinkToFit(true);//缩小以适合
    /**
     * 图片模块
     */

    $img = $cpsp2[$lan][0]["remarkimg2"];
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

        $drawing->setWidth($width>120 ? 120:$width);
        //$drawing->setHeight($height>130 ? 130:$height);
//$drawing->setHeight(150);


        //$drawing->setCoordinates($cola.'2');
        $drawing->setCoordinates($cola.'2');
        $drawing->setOffsetX(5);
        $drawing->setOffsetY(5);
        $drawing->setWorksheet($spreadsheet->getActiveSheet());
    }
    /* 图片模块 */


$thiscol = 15;  //当前行
    $itemcol = (count($itemarr)/3);
    if($itemcol != 0){

        for($l= $thiscol,$t= 0;$l< ($thiscol + $itemcol);$l++ ,$t++){


            $spreadsheet->getActiveSheet()->getStyle("{$cola}{$l}")->applyFromArray($styleArray1);
            $spreadsheet->getActiveSheet()->getStyle("{$colb}{$l}")->applyFromArray($styleArray1);
            $spreadsheet->getActiveSheet()->getStyle("{$colc}{$l}")->applyFromArray($styleArray1);

            $a = 0 + ($t * 3);
            $b = 1 + ($t * 3);
            $c = 2 + ($t * 3);

            $spreadsheet->getActiveSheet()->setCellValue("{$cola}{$l}", $itemarr[$a]);
            $spreadsheet->getActiveSheet()->setCellValue("{$colb}{$l}", $itemarr[$b]);
            $spreadsheet->getActiveSheet()->setCellValue("{$colc}{$l}", $itemarr[$c]);
            //$spreadsheet->getActiveSheet()->getStyle("{$cola}{$l}:{$colc}{$l}")->getAlignment()->setShrinkToFit(true);//缩小以适合

        }
    }else{
        $itemcol = 0;
    }




    $thiscol = 15 + $maxitem ;//当前行

        for($l= 15;$l < $thiscol;$l++) {  //填写样式


            $spreadsheet->getActiveSheet()->getStyle("{$cola}{$l}")->applyFromArray($styleArray1);
            $spreadsheet->getActiveSheet()->getStyle("{$colb}{$l}")->applyFromArray($styleArray1);
            $spreadsheet->getActiveSheet()->getStyle("{$colc}{$l}")->applyFromArray($styleArray1);
        }

    $rearr = array('Remarks','Price');
    $itemcol = count($rearr);

    for($l= $thiscol,$t= 0,$v = 5;$l< ($thiscol + 2);$l++ ,$t++,$v++){


        $spreadsheet->getActiveSheet()->getStyle("{$cola}{$l}")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->getStyle("{$colb}{$l}")->applyFromArray($styleArray1);
        $spreadsheet->getActiveSheet()->getStyle("{$colc}{$l}")->applyFromArray($styleArray1);


        $b = 0 + ($v * 2);
        $c = 1 + ($v * 2);

        $spreadsheet->getActiveSheet()->setCellValue("{$cola}{$l}", $rearr[$t]);
        $spreadsheet->getActiveSheet()->setCellValue("{$colb}{$l}", $fabarr[$b]);
        $spreadsheet->getActiveSheet()->setCellValue("{$colc}{$l}", $fabarr[$c]);

    }




} //1st for


} //$page;



//$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页
$spreadsheet->setActiveSheetIndex(0);  //設置當前活動表

unset($_SESSION['cpsall'] ); //注销SESSION

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
