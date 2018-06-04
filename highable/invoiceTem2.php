<?php
session_start();
header("Content-type: text/html; charset=utf-8");

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

$intem2 =  $_SESSION['invoiceTem2'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/invoiceTem2.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
//$spreadsheet->getActiveSheet()->getDefaultColumnDimension()->setWidth(20);  //设置默认列宽
$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(9);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(9);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(13);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(13);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(16);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(16);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(13);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('H')->setWidth(15);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('I')->setWidth(15);  //列宽度
$spreadsheet->getDefaultStyle()->getFont()->setSize(7);


$styleArray1 = [

    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
        'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
        'wrapText' => true,
        'ShrinkToFit'=>true,
    ],
    'font' => [
        'Size' => '8',
    ],

];

$styleArrayr = [

    'borders' => [

        'right' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],

    ],

];

$styleArraybu = [

    'borders' => [

        'bottom' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],

    ],

];

//填数据

$spreadsheet->getActiveSheet()->setCellValue('G5', $intem2["invoicedata"]["invoiceNumber"]);
$spreadsheet->getActiveSheet()->getStyle('G5')->applyFromArray($styleArray1);
$spreadsheet->getActiveSheet()->setCellValue('K11', $intem2["invoicedate"]);
$spreadsheet->getActiveSheet()->getStyle('K11')->applyFromArray($styleArray1);

$nowcol = 7;
$spreadsheet->getActiveSheet()->mergeCells("B{$nowcol}:E{$nowcol}");
$spreadsheet->getActiveSheet()->setCellValue('B'.$nowcol, $intem2["tosb"]);
$nowcol++;
for($i= 1,$l = 1; $i <= 4 ; $i++,$l++){
    $spreadsheet->getActiveSheet()->mergeCells("B{$nowcol}:E{$nowcol}");

    $spreadsheet->getActiveSheet()->getCell('B'.$nowcol)->setValue($intem2["invoicedata"]['a'.$l]);
    $spreadsheet->getActiveSheet()->getStyle('B'.$nowcol)->applyFromArray($styleArray1);
    $nowcol++;
}

$spreadsheet->getActiveSheet()->setCellValue('A14', $intem2["invoiceform"]["b1"]);
$spreadsheet->getActiveSheet()->setCellValue('B14', $intem2["invoiceform"]["b2"]);

$spreadsheet->getActiveSheet()->setCellValue('H14', $intem2["invoiceform"]["b3"]);
$spreadsheet->getActiveSheet()->setCellValue('I14', $intem2["invoiceform"]["b4"]);

$spreadsheet->getActiveSheet()->setCellValue('C15', 'SHIPMENT');
$spreadsheet->getActiveSheet()->setCellValue('D15', $intem2["invoiceform"]["b5"]);
$spreadsheet->getActiveSheet()->setCellValue('E15', 'TO');
$spreadsheet->getActiveSheet()->setCellValue('F15', $intem2["invoiceform"]["b6"]);

$spreadsheet->getActiveSheet()->setCellValue('C16', 'STYLE NO.');
$spreadsheet->getActiveSheet()->setCellValue('D16', 'FABRIC');
$spreadsheet->getActiveSheet()->setCellValue('E16', 'COLOUR+DESCRIPTION');
$spreadsheet->getActiveSheet()->setCellValue('G16', 'SHIP DATE');

$nowcol = 18;

////中部form
//$nowcol = 21;
//$spreadsheet->getActiveSheet()->setCellValue('H'.$nowcol, $intem2["invoiceform"]["price"]);
//$spreadsheet->getActiveSheet()->setCellValue('I'.$nowcol, $intem2["invoiceform"]["amout"]);
//
//$nowcol++;
//
for($x = 0 ,$c = 1; $c <= $intem2["invoiceform"]["formnum"]; $x++ ,$c++){

$f19 = 18 + 1 * $x;

$formarr = array('A'.$f19,'B'.$f19,'C'.$f19,'D'.$f19,'E'.$f19,'F'.$f19,'G'.$f19,'H'.$f19,'I'.$f19);

    for($i = 7,$y = 0; $i <= $intem2["invoiceform"]["brrnum"] ; $i++ ,$y++){

        $sheet->setCellValue($formarr[$y],  $intem2["invoiceform"]['b'.$i][$x]);

    }


    $nowcol = 22 +   1 * $c;



    if($x >12){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($nowcol, 1);
    }

}
//$nowcol = $intem2["invoiceform"]["formnum"] > 12 ? ($nowcol + 1) : 36;
////$spreadsheet->getActiveSheet()->getCell('A1')->setValue($nowcol);
//
//$spreadsheet->getActiveSheet()->setCellValue('A'.$nowcol, $intem2["invoiceform"]["coltb"]);
//$spreadsheet->getActiveSheet()->setCellValue('I'.$nowcol, $intem2["invoiceform"]["coltc"]);
//
//$nowcol++;
//$nowcol++;
//$spreadsheet->getActiveSheet()->getCell('D'.$nowcol)->setValue($intem2["invoiceform"]["formremark"]);
//$spreadsheet->getActiveSheet()->getStyle('D'.$nowcol)->applyFromArray($styleArray1);
//
//$nowcol++;
//$nowcol++;
//$nowcol++;
//$spreadsheet->getActiveSheet()->mergeCells("E{$nowcol}:H{$nowcol}");
//$spreadsheet->getActiveSheet()->getCell('E'.$nowcol)->setValue($intem2["remark"]["bottomremark"]);
//$spreadsheet->getActiveSheet()->getStyle("F{$nowcol}:H{$nowcol}")->applyFromArray($styleArray1);
//$nowcol++;
//$nowcol++;
//
//
//for($b = 1 ; $b<= $intem2["remark"]["crrnum"] ; $b++ ){
//    $spreadsheet->getActiveSheet()->mergeCells("F{$nowcol}:H{$nowcol}");
//    $spreadsheet->getActiveSheet()->getCell('F'.$nowcol)->setValue($intem2["remark"]["c".$b]);
//    $spreadsheet->getActiveSheet()->getStyle("F{$nowcol}:H{$nowcol}")->applyFromArray($styleArray1);
//    $nowcol++;
//
//
//}
//


////边栏样式
//$spreadsheet->getActiveSheet()->getStyle("A19:A{$nowcol}")->applyFromArray($styleArrayl);
//$spreadsheet->getActiveSheet()->getStyle("H19:H{$nowcol}")->applyFromArray($styleArrayr);
//$spreadsheet->getActiveSheet()->getStyle("A{$nowcol}:H{$nowcol}")->applyFromArray($styleArraybu);




$spreadsheet->getActiveSheet()->getPageSetup()
    ->setOrientation(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE); //打印橫向
$spreadsheet->getActiveSheet()->getPageSetup()
    ->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);//打印橫向 A4
$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

//unset($_SESSION['shipp1'] ); //注销SESSION

$output=  ($_GET['action'] == 'formdown' )? 1:0;
$nt = date("YmdHis",time()); //转换为日期。
$filenameout = 'intem2out'.$nt.'.xlsx';
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

