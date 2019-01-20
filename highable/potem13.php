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

$potem13 =  $_SESSION['potem13'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/potem13.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$spreadsheet->getActiveSheet()->setTitle("sheet1");
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
//$spreadsheet->getActiveSheet()->getDefaultColumnDimension()->setWidth(30);  //设置默认列宽
for($j=1;$j<=25;$j++){
    $col = chr(65 + $j);
    $spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(4);  //列宽度
}
for($j=0;$j<=10;$j++){
    $col = chr(65 + $j);
    $spreadsheet->getActiveSheet()->getColumnDimension('A'.$col)->setWidth(4);  //列宽度
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
$spreadsheet->getActiveSheet()->setCellValue('H9', 'TO: '.$potem13["tosb"]);
//$spreadsheet->getActiveSheet()->setCellValue('G11', 'DATE: '.$potem13["podate"]);
//
//$spreadsheet->getActiveSheet()->setCellValue('B16', $potem13["toaddr"]["a1"]);
//$spreadsheet->getActiveSheet()->setCellValue('B17', $potem13["toaddr"]["a2"]);
//$spreadsheet->getActiveSheet()->setCellValue('B18', $potem13["toaddr"]["a3"]);
$toaddr = array('Z9','B11','T11','B13','T13','F15','O15','Y15','B23','B34','U23','B41','K41','U41','Z41','AE41','J43','AC43');  //,'C12','D12','E12','F12','G12','H12','I12','J13','K12','K13','J5'

for($i = 1,$y = 0; $i <= count($toaddr) ; $i++ ,$y++){

    $sheet->setCellValue($toaddr[$y],  $potem13["toaddr"]["a".$i]);

}


//中部form

//$nowcol = 35;
////$spreadsheet->getActiveSheet()->mergeCells("A{$nowcol}:F{$nowcol}");
//$spreadsheet->getActiveSheet()->setCellValue('A'.$nowcol, 'PO NO: '.$potem13["orderform"]["midpono"].'   注：請在開發票時把“PONO”寫上，不可重復，并且寫上制單號）');
////$spreadsheet->getActiveSheet()->setCellValue('I'.$nowcol, $potem13["invoiceform"]["amout"]);


for($x = 0 ,$c = 1; $c <= $potem13["orderform"]["formnum"]; $x++ ,$c++){

$f19 = 44 + 1 * $x;




$formarr = array('B'.$f19,'E'.$f19,'I'.$f19,'N'.$f19,'X'.$f19,'AC'.$f19,);

    for($i = 1,$y = 0; $i <= $potem13["orderform"]["brrnum"] ; $i++ ,$y++){

        $sheet->setCellValue($formarr[$y],  $potem13["orderform"]['b'.$i][$x]);

    }

    $nowcol = 44  +   1 * $c;


    if($x >4){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($nowcol, 1);
    }
    $spreadsheet->getActiveSheet()->mergeCells("B{$nowcol}:D{$nowcol}");
    $spreadsheet->getActiveSheet()->mergeCells("E{$nowcol}:H{$nowcol}");
    $spreadsheet->getActiveSheet()->mergeCells("I{$nowcol}:M{$nowcol}");
    $spreadsheet->getActiveSheet()->mergeCells("N{$nowcol}:W{$nowcol}");
    $spreadsheet->getActiveSheet()->mergeCells("X{$nowcol}:AB{$nowcol}");
    $spreadsheet->getActiveSheet()->mergeCells("AC{$nowcol}:AJ{$nowcol}");
}
$nowcol = $potem13["orderform"]["formnum"] > 4 ? ($nowcol + 1) : 50;
$spreadsheet->getActiveSheet()->setCellValue('Q'.$nowcol, $potem13["toaddr"]["a19"]);
//////$spreadsheet->getActiveSheet()->mergeCells("A{$nowcol}:E{$nowcol}");
$spreadsheet->getActiveSheet()->setCellValue('J18', $potem13["toaddr"]["a20"]);
$nowcol++;
$nowcol++;


$spreadsheet->getActiveSheet()->setCellValue('O'.$nowcol, $potem13["remark"]["c1"]);
$nowcol++;
$spreadsheet->getActiveSheet()->setCellValue('O'.$nowcol, $potem13["remark"]["c2"]);
$nowcol++;
//$nowcol++;
//
//$spreadsheet->getActiveSheet()->setCellValue('B'.$nowcol, 'E-MAIL (電郵): '.$potem13["remark"]["c3"]);
//$spreadsheet->getActiveSheet()->setCellValue('I'.$nowcol, $potem13["remark"]["c4"]);
////$spreadsheet->getActiveSheet()->mergeCells("A{$nowcol}:E{$nowcol}");
$spreadsheet->getActiveSheet()->setCellValue('O'.$nowcol, $potem13["remark"]["c3"]);
$nowcol++;

$spreadsheet->getActiveSheet()->setCellValue('O'.$nowcol, $potem13["remark"]["c4"]);

$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

unset($_SESSION['potem13'] ); //注销SESSION

$output=  ($_GET['action'] == 'formdown' )? 1:0;
$nt = date("YmdHis",time()); //转换为日期。
$filenameout = 'potem13out'.$nt.'.xlsx';
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

