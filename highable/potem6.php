<?php
session_start();
header("Content-type: text/html; charset=utf-8");
/*嘉和信*/

require_once('autoloadconfig.php');  //判断是否在线

require_once ('img.php');

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;

$potem6 =  $_SESSION['potem6'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/potem6.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$spreadsheet->getActiveSheet()->setTitle("sheet1");
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
//$spreadsheet->getActiveSheet()->getDefaultColumnDimension()->setWidth(20);  //设置默认列宽
$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(25);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(25);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(25);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(20);  //列宽度
//
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


//填数据
$spreadsheet->getActiveSheet()->setCellValue('B6', $potem6["tosb"]);
$spreadsheet->getActiveSheet()->setCellValue('F7', $potem6 ["podate"]);
$spreadsheet->getActiveSheet()->setCellValue('F6', $potem6["toaddr"]["a1"]);
$spreadsheet->getActiveSheet()->setCellValue('B7', $potem6["toaddr"]["a2"]);



//中部form
//$spreadsheet->getActiveSheet()->setCellValue('B11', $potem6["toaddr"]["a7"]);
//$spreadsheet->getActiveSheet()->setCellValue('B12', $potem6["orderform"]["midpono"]);

$nowcol = 11;
//$spreadsheet->getActiveSheet()->mergeCells("A{$nowcol}:F{$nowcol}");
//$spreadsheet->getActiveSheet()->setCellValue('A'.$nowcol, '(PO NO:  '.$potem6["orderform"]["midpono"].' 注：請在開發票時把"PO NO"寫上，不可重複)');
////$spreadsheet->getActiveSheet()->setCellValue('I'.$nowcol, $potem6["invoiceform"]["amout"]);
////
//$nowcol++;
//$nowcol++;
//
for($x = 0 ,$c = 1; $x <= $potem6["orderform"]["formnum"]; $x++ ,$c++){

$f19 = 11 + 1 * $x;

$spreadsheet->getActiveSheet()->mergeCells("B{$f19}:D{$f19}");

$formarr = array('A'.$f19,'B'.$f19,'E'.$f19,'F'.$f19,'G'.$f19);

    for($i = 1,$y = 0; $i <= $potem6["orderform"]["brrnum"] ; $i++ ,$y++){

        $sheet->setCellValue($formarr[$y],  $potem6["orderform"]['b'.$i][$x]);

    }


    $nowcol = 11  +   1 * $c;



    if($x >9){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($nowcol, 1);
    }

}



//底部REMARK
$nowcol = $potem6["orderform"]["formnum"] > 9? ($nowcol + 1) : 22;
//$spreadsheet->getActiveSheet()->getCell('A1')->setValue($nowcol);

$spreadsheet->getActiveSheet()->setCellValue('E'.$nowcol, $potem6["toaddr"]["a3"]);
$spreadsheet->getActiveSheet()->setCellValue('F'.$nowcol, $potem6["toaddr"]["a4"]);
$spreadsheet->getActiveSheet()->setCellValue('G'.$nowcol, $potem6["toaddr"]["a5"]);
$nowcol++;

//$spreadsheet->getActiveSheet()->mergeCells("A{$nowcol}:E{$nowcol}");
$spreadsheet->getActiveSheet()->setCellValue('A'.$nowcol, '備註：'.$potem6["remark"]["c1"]);
$nowcol++;
////
////$spreadsheet->getActiveSheet()->mergeCells("A{$nowcol}:E{$nowcol}");
//$spreadsheet->getActiveSheet()->setCellValue('B'.$nowcol, $potem6["remark"]["c1"]);
$nowcol++;
$nowcol++;
//
////
//$spreadsheet->getActiveSheet()->mergeCells("A{$nowcol}:E{$nowcol}");
$spreadsheet->getActiveSheet()->setCellValue('A'.$nowcol, $potem6["remark"]["c2"]);
$nowcol++;
//
////$spreadsheet->getActiveSheet()->mergeCells("A{$nowcol}:E{$nowcol}");
$spreadsheet->getActiveSheet()->setCellValue('A'.$nowcol, 'FAX:'.$potem6["remark"]["c3"]);
$nowcol++;
$spreadsheet->getActiveSheet()->setCellValue('A'.$nowcol, 'E-mail:'.$potem6["remark"]["c4"]);
$nowcol++;
$spreadsheet->getActiveSheet()->setCellValue('A'.$nowcol, '交貨期:'.$potem6["remark"]["c5"]);


$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

unset($_SESSION['potem6'] ); //注销SESSION

$output=  ($_GET['action'] == 'formdown' )? 1:0;
$nt = date("YmdHis",time()); //转换为日期。
$filenameout = 'potem6out'.$nt.'.xlsx';
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

