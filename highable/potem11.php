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

$potem11 =  $_SESSION['potem11'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/potem11.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$spreadsheet->getActiveSheet()->setTitle("sheet1");
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
//$spreadsheet->getActiveSheet()->getDefaultColumnDimension()->setWidth(20);  //设置默认列宽
//$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(25);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(5);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(15);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(30);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(20);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(15);  //列宽度
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
$spreadsheet->getActiveSheet()->setCellValue('B11', 'TO: '.$potem11["tosb"]);
$spreadsheet->getActiveSheet()->setCellValue('G11', 'DATE: '.$potem11["podate"]);

$spreadsheet->getActiveSheet()->setCellValue('B16', $potem11["toaddr"]["a1"]);
$spreadsheet->getActiveSheet()->setCellValue('B17', $potem11["toaddr"]["a2"]);
$spreadsheet->getActiveSheet()->setCellValue('B18', $potem11["toaddr"]["a3"]);
$spreadsheet->getActiveSheet()->setCellValue('B19', $potem11["toaddr"]["a4"]);
$spreadsheet->getActiveSheet()->setCellValue('B21', 'Contact (聯絡人) :'.$potem11["toaddr"]["a5"]);
$spreadsheet->getActiveSheet()->setCellValue('B23', 'Tel (電話): '.$potem11["toaddr"]["a6"]);
$spreadsheet->getActiveSheet()->setCellValue('B25', 'Fax (傳真): '.$potem11["toaddr"]["a7"]);
$spreadsheet->getActiveSheet()->setCellValue('B27', 'Email (電郵): '.$potem11["toaddr"]["a8"]);
$spreadsheet->getActiveSheet()->setCellValue('B29', 'Customer PO# (客戶訂單號碼):'.$potem11["toaddr"]["a9"]);
$spreadsheet->getActiveSheet()->setCellValue('B31', 'Payment currency (付款幣值):'.$potem11["toaddr"]["a10"]);


$spreadsheet->getActiveSheet()->setCellValue('G16', $potem11["toaddr"]["a11"]);
$spreadsheet->getActiveSheet()->setCellValue('G17', $potem11["toaddr"]["a12"]);
$spreadsheet->getActiveSheet()->setCellValue('G18', $potem11["toaddr"]["a13"]);
$spreadsheet->getActiveSheet()->setCellValue('G19', $potem11["toaddr"]["a14"]);
$spreadsheet->getActiveSheet()->setCellValue('G21', 'Contact (聯絡人) :'.$potem11["toaddr"]["a15"]);
$spreadsheet->getActiveSheet()->setCellValue('G23', 'Tel (電話): '.$potem11["toaddr"]["a16"]);
$spreadsheet->getActiveSheet()->setCellValue('G25', 'Fax (傳真): '.$potem11["toaddr"]["a17"]);
$spreadsheet->getActiveSheet()->setCellValue('G27', 'Ship mode (運輸方式) :'.$potem11["toaddr"]["a18"]);
$spreadsheet->getActiveSheet()->setCellValue('G29', 'Port of Discharge (港口名稱): '.$potem11["toaddr"]["a19"]);


//中部form

//$nowcol = 35;
////$spreadsheet->getActiveSheet()->mergeCells("A{$nowcol}:F{$nowcol}");
//$spreadsheet->getActiveSheet()->setCellValue('A'.$nowcol, 'PO NO: '.$potem11["orderform"]["midpono"].'   注：請在開發票時把“PONO”寫上，不可重復，并且寫上制單號）');
////$spreadsheet->getActiveSheet()->setCellValue('I'.$nowcol, $potem11["invoiceform"]["amout"]);


for($x = 0 ,$c = 1; $c <= $potem11["orderform"]["formnum"]; $x++ ,$c++){

$f19 = 35 + 1 * $x;

$spreadsheet->getActiveSheet()->mergeCells("B{$f19}:C{$f19}");
$spreadsheet->getActiveSheet()->mergeCells("F{$f19}:G{$f19}");

$formarr = array('B'.$f19,'D'.$f19,'E'.$f19,'F'.$f19,'H'.$f19,'I'.$f19);

    for($i = 1,$y = 0; $i <= $potem11["orderform"]["brrnum"] ; $i++ ,$y++){

        $sheet->setCellValue($formarr[$y],  $potem11["orderform"]['b'.$i][$x]);

    }

    $nowcol = 35  +   1 * $c;


    if($x >1){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($nowcol, 1);
    }

}
$nowcol = $potem11["orderform"]["formnum"] > 1 ? ($nowcol + 2) : 39;
////$spreadsheet->getActiveSheet()->getCell('A1')->setValue($nowcol); 貨送以下地址
////$spreadsheet->getActiveSheet()->mergeCells("A{$nowcol}:E{$nowcol}");
////$spreadsheet->getActiveSheet()->setCellValue('A'.$nowcol, '貨送以下地址');
////$nowcol++;
//
//$spreadsheet->getActiveSheet()->mergeCells("A{$nowcol}:E{$nowcol}");
$spreadsheet->getActiveSheet()->setCellValue('D'.$nowcol, $potem11["toaddr"]["a20"]);
$nowcol++;
$nowcol++;
$nowcol++;

////$spreadsheet->getActiveSheet()->mergeCells("A{$nowcol}:E{$nowcol}");
$spreadsheet->getActiveSheet()->setCellValue('D'.$nowcol, $potem11["toaddr"]["a21"]);
$nowcol++;
$nowcol++;
$nowcol++;
$nowcol++;



$spreadsheet->getActiveSheet()->setCellValue('B'.$nowcol, 'ORDERED BY (經手人) :'.$potem11["remark"]["c1"]);
$spreadsheet->getActiveSheet()->setCellValue('I'.$nowcol, $potem11["remark"]["c2"]);
$nowcol++;
$nowcol++;

$spreadsheet->getActiveSheet()->setCellValue('B'.$nowcol, 'E-MAIL (電郵): '.$potem11["remark"]["c3"]);
$spreadsheet->getActiveSheet()->setCellValue('I'.$nowcol, $potem11["remark"]["c4"]);
////$spreadsheet->getActiveSheet()->mergeCells("A{$nowcol}:E{$nowcol}");
//$spreadsheet->getActiveSheet()->setCellValue('E'.$nowcol, $potem11["remark"]["c3"]);
//$nowcol++;
//
//$spreadsheet->getActiveSheet()->setCellValue('A'.$nowcol, $potem11["remark"]["c4"]);
//
$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页
//
//unset($_SESSION['potem11'] ); //注销SESSION

$output=  ($_GET['action'] == 'formdown' )? 1:0;
$nt = date("YmdHis",time()); //转换为日期。
$filenameout = 'potem11out'.$nt.'.xlsx';
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

