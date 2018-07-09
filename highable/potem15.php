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

$potem15 =  $_SESSION['potem15'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/potem15.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$spreadsheet->getActiveSheet()->setTitle("sheet1");
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
//$spreadsheet->getActiveSheet()->getDefaultColumnDimension()->setWidth(30);  //设置默认列宽
//for($j=0;$j<=6;$j++){
//    $col = chr(65 + $j);
//    $spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(20);  //列宽度
//}

$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(25);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(15);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(10);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(15);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('H')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('I')->setWidth(30);  //列宽度

$spreadsheet->getDefaultStyle()->getFont()->setSize(7);

$styleArray1 = [
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
        'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
        'wrapText' => true,
        'ShrinkToFit'=>true,
    ],
    'font' => [
        'Size' => '6',
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

    ]

];
$styleArray2 = [
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
        'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
        'wrapText' => true,
        'ShrinkToFit'=>true,
    ],
    'font' => [
        'Size' => '5',
    ]

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
$spreadsheet->getActiveSheet()->setCellValue('B6', $potem15["tosb"]);
$spreadsheet->getActiveSheet()->setCellValue('B8', $potem15["podate"]);
////
//$spreadsheet->getActiveSheet()->setCellValue('G8', $potem15["toaddr"]["a1"]);
//$spreadsheet->getActiveSheet()->setCellValue('B9', $potem15["toaddr"]["a2"]);
//$spreadsheet->getActiveSheet()->setCellValue('B11', $potem15["toaddr"]["a3"]);
//$spreadsheet->getActiveSheet()->setCellValue('B12', $potem15["toaddr"]["a4"]);
$toaddr = array('H6','B7','H7','H8','B9','H9','B12','B13','B14','B15','B16','B17','B18','B19','B20','G12','G13','G14','G15','G16','G17','G18','G19','G20');  //

for($i = 1,$y = 0; $i <= count($toaddr) ; $i++ ,$y++){

    $sheet->setCellValue($toaddr[$y],  $potem15["toaddr"]["a".$i]);

}
//
//
////中部form
//
//$nowcol = 14;
////////$spreadsheet->getActiveSheet()->mergeCells("A{$nowcol}:F{$nowcol}");
//$spreadsheet->getActiveSheet()->setCellValue('B'.$nowcol, $potem15["orderform"]["midpono"]);
////////$spreadsheet->getActiveSheet()->setCellValue('I'.$nowcol, $potem15["invoiceform"]["amout"]);
////
////
//for($x = 0 ,$c = 1; $c <= $potem15["orderform"]["formnum"]; $x++ ,$c++){
//
//$f19 = 15 + 1 * $x;
//
//    $spreadsheet->getActiveSheet()->mergeCells("A{$f19}:B{$f19}");
//    $spreadsheet->getActiveSheet()->mergeCells("C{$f19}:G{$f19}");
//
//
//$formarr = array('A'.$f19,'C'.$f19);
//
//    for($i = 1,$y = 0; $i <= $potem15["orderform"]["brrnum"] ; $i++ ,$y++){
//
//        $sheet->setCellValue($formarr[$y],  $potem15["orderform"]['b'.$i][$x]);
//
//    }
//
//    $nowcol = 15  +  1 * $c;
//
//
//    if($x >4){
//        $spreadsheet->getActiveSheet()->insertNewRowBefore($nowcol, 1);
//    }
//
//}
//$nowcol = $potem15["orderform"]["formnum"] > 4 ? ($nowcol + 1) : 21;
//$spreadsheet->getActiveSheet()->setCellValue('C'.$nowcol, $potem15["toaddr"]["a5"]);
//$nowcol++;
//
//$spreadsheet->getActiveSheet()->setCellValue('C'.$nowcol, $potem15["toaddr"]["a6"]);
//$nowcol++;
//$spreadsheet->getActiveSheet()->setCellValue('C'.$nowcol, $potem15["toaddr"]["a7"]);
//$nowcol++;
//$spreadsheet->getActiveSheet()->setCellValue('A'.$nowcol, $potem15["toaddr"]["a8"]);
//$nowcol++;
//$nowcol++;
//$nowcol++;
//$nowcol++;
//$nowcol++;
//$nowcol++;
//$nowcol++;
//
//$spreadsheet->getActiveSheet()->setCellValue('C'.$nowcol, $potem15["remark"]["c1"]);
//$nowcol++;
//$spreadsheet->getActiveSheet()->setCellValue('C'.$nowcol, $potem15["remark"]["c2"]);

////
////$spreadsheet->getActiveSheet()->setCellValue('B'.$nowcol, 'E-MAIL (電郵): '.$potem15["remark"]["c3"]);
////$spreadsheet->getActiveSheet()->setCellValue('I'.$nowcol, $potem15["remark"]["c4"]);
//////$spreadsheet->getActiveSheet()->mergeCells("A{$nowcol}:E{$nowcol}");
//$spreadsheet->getActiveSheet()->setCellValue('O'.$nowcol, $potem15["remark"]["c3"]);
//$nowcol++;
//
//$spreadsheet->getActiveSheet()->setCellValue('O'.$nowcol, $potem15["remark"]["c4"]);

$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

unset($_SESSION['potem15'] ); //注销SESSION

$output=  ($_GET['action'] == 'formdown' )? 1:0;
$nt = date("YmdHis",time()); //转换为日期。
$filenameout = 'potem15out'.$nt.'.xlsx';
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

