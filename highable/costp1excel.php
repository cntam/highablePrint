<?php
session_start();

$costp1 =  $_SESSION['costp1'];
require '/home/pan/vendor/autoload.php';

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

    $img = $costp1['remarkimg2'];

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
    $drawing->setName($costp1['costname']);
    $drawing->setDescription($costp1['costname']);
//$drawing->setImageResource($gdImage);
    $drawing->setImageResource($img);
    $drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
    $drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
//$drawing->setHeight($width);

//$drawing->setHeight($width>550 ? 550:$width);
//$drawing->setWidth($width>500 ? 500:$width);

    $resw = $width < 520 ? 0 : 2;
    $resh = $height < 490 ? 0 : 3;
    $res = $resw + $resh;
    switch ($res)
    {
        case "2":
            $drawing->setWidth(520);
            break;
        case "3":
            $drawing->setHeight(490);
            break;

        default:
            $drawing->setHeight($height>550 ? 490:$height);
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
    $spreadsheet->getActiveSheet()->setCellValue('B28', $costp1['costname']);
    $spreadsheet->getActiveSheet()->setCellValue('A29', $costp1['costno']);
    $spreadsheet->getActiveSheet()->setCellValue('B29', $costp1['ccno2']);
    $spreadsheet->getActiveSheet()->setCellValue('E31', $costp1['costno']);

    $spreadsheet->getActiveSheet()->setCellValue('B31', $costp1['fab']['a1']);
    $spreadsheet->getActiveSheet()->setCellValue('B32', $costp1['fab']['b1']);
    $spreadsheet->getActiveSheet()->setCellValue('B33', $costp1['fab']['c1']);
    $spreadsheet->getActiveSheet()->setCellValue('B34', $costp1['fab']['d1']);
    $spreadsheet->getActiveSheet()->setCellValue('B35', $costp1['fab']['e1']);
    $spreadsheet->getActiveSheet()->setCellValue('B36', $costp1['fab']['f1']);
    $spreadsheet->getActiveSheet()->setCellValue('B37', $costp1['fab']['g1']);
    $spreadsheet->getActiveSheet()->setCellValue('B38', $costp1['fab']['h1']);

    $spreadsheet->getActiveSheet()->setCellValue('E33', $costp1['fab']['c2']);
    $spreadsheet->getActiveSheet()->setCellValue('E34', $costp1['fab']['d2']);
    $spreadsheet->getActiveSheet()->setCellValue('E35', $costp1['fab']['e2']);
    $spreadsheet->getActiveSheet()->setCellValue('E36', $costp1['fab']['f2']);
    $spreadsheet->getActiveSheet()->setCellValue('E37', $costp1['fab']['g2']);
    $spreadsheet->getActiveSheet()->setCellValue('E38', $costp1['fab']['h2']);

    $spreadsheet->getActiveSheet()->setCellValue('H33', $costp1['fab']['c3']);
    $spreadsheet->getActiveSheet()->setCellValue('H34', $costp1['fab']['d3']);
    $spreadsheet->getActiveSheet()->setCellValue('H35', $costp1['fab']['e3']);
    $spreadsheet->getActiveSheet()->setCellValue('H36', $costp1['fab']['f3']);
    $spreadsheet->getActiveSheet()->setCellValue('H37', $costp1['fab']['g3']);




    unset($_SESSION['costp1'] ); //注销SESSION

// Set active sheet index to the first sheet, so Excel opens this as the first sheet
    $spreadsheet->setActiveSheetIndex(0); //返回第一页

    $spreadsheet->getActiveSheet()->getPageMargins()->setRight(0.1); //设置打印边距
    $spreadsheet->getActiveSheet()->getPageMargins()->setLeft(0.1); //

$output=  ($_GET['action'] == 'formdown' )? 1:0;
//$output= 0;
$nt = date("YmdHis",time()); //转换为日期。
$filenameout = 'costp1out'.$nt.'.xlsx';
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
