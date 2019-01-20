<?php
session_start();
header("Content-type: text/html; charset=utf-8");

require_once('autoloadconfig.php');  //判断是否在线

require_once ('img.php');

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;

$potem14 =  $_SESSION['potem14'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/potem14.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$spreadsheet->getActiveSheet()->setTitle("sheet1");
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
//$spreadsheet->getActiveSheet()->getDefaultColumnDimension()->setWidth(30);  //设置默认列宽
for($j=0;$j<=6;$j++){
    $col = chr(65 + $j);
    $spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(20);  //列宽度
}

//$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(25);  //列宽度
//$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(5);  //列宽度
//$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(15);  //列宽度

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
$spreadsheet->getActiveSheet()->setCellValue('B8', 'TO: '.$potem14["tosb"]);
$spreadsheet->getActiveSheet()->setCellValue('B10', 'DATE: '.$potem14["podate"]);
////
$spreadsheet->getActiveSheet()->setCellValue('G8', $potem14["toaddr"]["a1"]);
$spreadsheet->getActiveSheet()->setCellValue('B9', $potem14["toaddr"]["a2"]);
$spreadsheet->getActiveSheet()->setCellValue('B11', $potem14["toaddr"]["a3"]);
$spreadsheet->getActiveSheet()->setCellValue('B12', $potem14["toaddr"]["a4"]);
//$toaddr = array('Z9','B11','T11','B13','T13','F15','O15','Y15','B23','B34','U23','B41','K41','U41','Z41','AE41','J43','AC43');  //,'C12','D12','E12','F12','G12','H12','I12','J13','K12','K13','J5'
//
//for($i = 1,$y = 0; $i <= count($toaddr) ; $i++ ,$y++){
//
//    $sheet->setCellValue($toaddr[$y],  $potem14["toaddr"]["a".$i]);
//
//}
//
//
////中部form
//
$nowcol = 14;
//////$spreadsheet->getActiveSheet()->mergeCells("A{$nowcol}:F{$nowcol}");
$spreadsheet->getActiveSheet()->setCellValue('B'.$nowcol, $potem14["orderform"]["midpono"]);
//////$spreadsheet->getActiveSheet()->setCellValue('I'.$nowcol, $potem14["invoiceform"]["amout"]);
//
//
for($x = 0 ,$c = 1; $c <= $potem14["orderform"]["formnum"]; $x++ ,$c++){

$f19 = 15 + 1 * $x;

    $spreadsheet->getActiveSheet()->mergeCells("A{$f19}:B{$f19}");
    $spreadsheet->getActiveSheet()->mergeCells("C{$f19}:G{$f19}");


$formarr = array('A'.$f19,'C'.$f19);

    for($i = 1,$y = 0; $i <= $potem14["orderform"]["brrnum"] ; $i++ ,$y++){

        $sheet->setCellValue($formarr[$y],  $potem14["orderform"]['b'.$i][$x]);

    }

    $nowcol = 15  +  1 * $c;


    if($x >4){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($nowcol, 1);
    }

}
$nowcol = $potem14["orderform"]["formnum"] > 4 ? ($nowcol + 1) : 21;
$spreadsheet->getActiveSheet()->setCellValue('C'.$nowcol, $potem14["toaddr"]["a5"]);
$nowcol++;

$spreadsheet->getActiveSheet()->setCellValue('C'.$nowcol, $potem14["toaddr"]["a6"]);
$nowcol++;
$spreadsheet->getActiveSheet()->setCellValue('C'.$nowcol, $potem14["toaddr"]["a7"]);
$nowcol++;
$spreadsheet->getActiveSheet()->setCellValue('A'.$nowcol, $potem14["toaddr"]["a8"]);
$nowcol++;
$nowcol++;
$nowcol++;
$nowcol++;
$nowcol++;
$nowcol++;
$nowcol++;

$spreadsheet->getActiveSheet()->setCellValue('C'.$nowcol, $potem14["remark"]["c1"]);
$nowcol++;
$spreadsheet->getActiveSheet()->setCellValue('C'.$nowcol, $potem14["remark"]["c2"]);

////
////$spreadsheet->getActiveSheet()->setCellValue('B'.$nowcol, 'E-MAIL (電郵): '.$potem14["remark"]["c3"]);
////$spreadsheet->getActiveSheet()->setCellValue('I'.$nowcol, $potem14["remark"]["c4"]);
//////$spreadsheet->getActiveSheet()->mergeCells("A{$nowcol}:E{$nowcol}");
//$spreadsheet->getActiveSheet()->setCellValue('O'.$nowcol, $potem14["remark"]["c3"]);
//$nowcol++;
//
//$spreadsheet->getActiveSheet()->setCellValue('O'.$nowcol, $potem14["remark"]["c4"]);

$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

unset($_SESSION['potem14'] ); //注销SESSION

$output=  ($_GET['action'] == 'formdown' )? 1:0;
$nt = date("YmdHis",time()); //转换为日期。
$filenameout = 'potem14out'.$nt.'.xlsx';
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

