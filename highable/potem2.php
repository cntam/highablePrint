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

$potem2 =  $_SESSION['potem2'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/potem2.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$spreadsheet->getActiveSheet()->setTitle("sheet1");
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
//$spreadsheet->getActiveSheet()->getDefaultColumnDimension()->setWidth(20);  //设置默认列宽
$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(16);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(16);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(16);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(16);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(16);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(16);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(16);  //列宽度
//$spreadsheet->getActiveSheet()->getColumnDimension('H')->setWidth(15);  //列宽度
//$spreadsheet->getActiveSheet()->getColumnDimension('I')->setWidth(15);  //列宽度
$spreadsheet->getDefaultStyle()->getFont()->setSize(7);

$styleArray1 = [
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
        'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
        'wrapText' => true,
        'ShrinkToFit'=>true,
    ],
    'font' => [
        'Size' => '8',
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
$spreadsheet->getActiveSheet()->setCellValue('B7', $potem2["tosb"]);
$spreadsheet->getActiveSheet()->setCellValue('G7', $potem2 ["podate"]);
$spreadsheet->getActiveSheet()->setCellValue('B8', $potem2["toaddr"]["a1"]);
$spreadsheet->getActiveSheet()->setCellValue('B9', $potem2["toaddr"]["a2"].'  FAX: '.$potem2["toaddr"]["a3"]);
$spreadsheet->getActiveSheet()->setCellValue('B10', $potem2["toaddr"]["a4"]);
$spreadsheet->getActiveSheet()->setCellValue('G10', $potem2["toaddr"]["a5"]);


//中部form

$nowcol = 15;
$spreadsheet->getActiveSheet()->mergeCells("A{$nowcol}:F{$nowcol}");
$spreadsheet->getActiveSheet()->setCellValue('A'.$nowcol, '(PO NO:  '.$potem2["orderform"]["midpono"].' 注：請在開發票時把"PO NO"寫上，不可重複)');
//$spreadsheet->getActiveSheet()->setCellValue('I'.$nowcol, $potem2["invoiceform"]["amout"]);
//
$nowcol++;
$nowcol++;

for($x = 0 ,$c = 1; $x <= $potem2["orderform"]["formnum"]; $x++ ,$c++){

$f19 = 17 + 1 * $x;

$spreadsheet->getActiveSheet()->mergeCells("B{$f19}:E{$f19}");

$formarr = array('A'.$f19,'B'.$f19,'F'.$f19,'G'.$f19);

    for($i = 1,$y = 0; $i <= $potem2["orderform"]["brrnum"] ; $i++ ,$y++){

        $sheet->setCellValue($formarr[$y],  $potem2["orderform"]['b'.$i][$x]);

    }


    $nowcol = 17  +   1 * $c;

    $spreadsheet->getActiveSheet()->getStyle('A'.$f19)->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->getStyle("B{$f19}:E{$f19}")->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->getStyle('F'.$f19)->applyFromArray($styleArray1);
    $spreadsheet->getActiveSheet()->getStyle('G'.$f19)->applyFromArray($styleArray1);

    if($x >6){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($nowcol, 1);
    }

}
$nowcol = $potem2["orderform"]["formnum"] > 6 ? ($nowcol + 2) : 27;
//$spreadsheet->getActiveSheet()->getCell('A1')->setValue($nowcol); 貨送以下地址
$spreadsheet->getActiveSheet()->mergeCells("A{$nowcol}:E{$nowcol}");
$spreadsheet->getActiveSheet()->setCellValue('A'.$nowcol, '貨送以下地址');
$nowcol++;

$spreadsheet->getActiveSheet()->mergeCells("A{$nowcol}:E{$nowcol}");
$spreadsheet->getActiveSheet()->setCellValue('A'.$nowcol, $potem2["remark"]["c1"]);
$nowcol++;


$spreadsheet->getActiveSheet()->mergeCells("A{$nowcol}:E{$nowcol}");
$spreadsheet->getActiveSheet()->setCellValue('A'.$nowcol, $potem2["remark"]["c2"]);
$nowcol++;
$nowcol++;
$spreadsheet->getActiveSheet()->mergeCells("A{$nowcol}:E{$nowcol}");
$spreadsheet->getActiveSheet()->setCellValue('A'.$nowcol, $potem2["remark"]["c3"]);



$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

//unset($_SESSION['potem2'] ); //注销SESSION

$output=  ($_GET['action'] == 'formdown' )? 1:0;
$nt = date("YmdHis",time()); //转换为日期。
$filenameout = 'potem2out'.$nt.'.xlsx';
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

