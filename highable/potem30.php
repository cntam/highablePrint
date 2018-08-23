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

$potem30 =  $_SESSION['potem30'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/potem30.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$spreadsheet->getActiveSheet()->setTitle("sheet1");
$spreadsheet->getDefaultStyle()->getFont()->setName('SimSun');
//$spreadsheet->getActiveSheet()->getDefaultColumnDimension()->setWidth(30);  //设置默认列宽
for($j=0;$j<=8;$j++){
    $col = chr(65 + $j);
    $spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(15);  //列宽度
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
$spreadsheet->getActiveSheet()->setCellValue('A9', $potem30["tosb"]);
$spreadsheet->getActiveSheet()->setCellValue('B18', $potem30["podate"]);

if(1 == $potem30["toaddr"]["a1"]){
  $titlecon = 'HIGH ABLE INVESTMENT LIMITED';
}elseif (2 == $potem30["toaddr"]["a1"]){
    $titlecon = 'IRONDALE FASHION INTERNATIONAL LIMITED';
}
$spreadsheet->getActiveSheet()->setCellValue('A1', $titlecon);
$spreadsheet->getActiveSheet()->setCellValue('G9', $potem30["toaddr"]["a2"]);
$spreadsheet->getActiveSheet()->setCellValue('A10', $potem30["toaddr"]["a3"]);
$spreadsheet->getActiveSheet()->setCellValue('A11', $potem30["toaddr"]["a4"]);
$spreadsheet->getActiveSheet()->setCellValue('A12', $potem30["toaddr"]["a5"]);
$spreadsheet->getActiveSheet()->setCellValue('A13', '电话：'.$potem30["toaddr"]["a6"]); //电话：0571-86312008 传真：0571-86312007
$spreadsheet->getActiveSheet()->setCellValue('A14', '传真：'.$potem30["toaddr"]["a7"]); //电话：0571-86312008 传真：0571-86312007
$spreadsheet->getActiveSheet()->setCellValue('A15', $potem30["toaddr"]["a8"]);

if(1 == $potem30["toaddr"]["a9"]){
    $amount = 'Amount(RMB)';
}elseif (2 == $potem30["toaddr"]["a9"]){
    $amount = 'Amount(HKD)';
}elseif (3 == $potem30["toaddr"]["a9"]){
    $amount = 'Amount(USD)';
}
$spreadsheet->getActiveSheet()->setCellValue('I19', $amount);
$spreadsheet->getActiveSheet()->setCellValue('A21', $potem30["orderform"]['b1'][0]);
$spreadsheet->getActiveSheet()->setCellValue('G21', $potem30["orderform"]['b2'][0]);
$spreadsheet->getActiveSheet()->setCellValue('I21', $potem30["orderform"]['b3'][0]);
$spreadsheet->getActiveSheet()->setCellValue('B22', $potem30["orderform"]['b4'][0]);
$spreadsheet->getActiveSheet()->setCellValue('G22', $potem30["orderform"]['b5'][0]);
$spreadsheet->getActiveSheet()->setCellValue('B23', $potem30["orderform"]['b6'][0]);

if(1 == $potem30["toaddr"]["a11"]){
    $um = 'U/M';
}elseif (2 == $potem30["toaddr"]["a11"]){
    $um = 'U/Y';
}
$spreadsheet->getActiveSheet()->setCellValue('H25', $um);
$spreadsheet->getActiveSheet()->setCellValue('B26', $potem30["toaddr"]["a12"]);
$spreadsheet->getActiveSheet()->setCellValue('F26', $potem30["toaddr"]["a13"]);
$spreadsheet->getActiveSheet()->setCellValue('G26', $potem30["toaddr"]["a14"]);
$spreadsheet->getActiveSheet()->setCellValue('H26', $potem30["toaddr"]["a15"]);
$spreadsheet->getActiveSheet()->setCellValue('G27', $potem30["toaddr"]["a16"]);
$spreadsheet->getActiveSheet()->setCellValue('H27', $potem30["toaddr"]["a17"]);

$spreadsheet->getActiveSheet()->setCellValue('A28', 'Total   Amount  ：'.$potem30["toaddr"]["a18"]);
$spreadsheet->getActiveSheet()->setCellValue('C28', $potem30["toaddr"]["a19"]);
$spreadsheet->getActiveSheet()->setCellValue('A29', 'Payment  Terms：'.$potem30["toaddr"]["a20"]);
$spreadsheet->getActiveSheet()->setCellValue('A30', 'Price   Terms    ：'.$potem30["toaddr"]["a21"]);


$spreadsheet->getActiveSheet()->setCellValue('B32', 'AMOUNT&QUANTITY WITHIN THE TOLERANCE OF '.$potem30["remark"]["c1"].' MORE OR LESS IS ONLY ALLOWED.');
$spreadsheet->getActiveSheet()->setCellValue('B36', 'YOU HAVE TO SUBMIT '.$potem30["remark"]["c2"].'SHIPMENT SAMPLE FOR OUR APPROVAL BEFORE'.$potem30["remark"]["c3"].'OF SHIPMENT.');
//YOU HAVE TO SUBMIT 4YDS SHIPMENT SAMPLE FOR OUR APPROVAL BEFORE 7 DAYS OF SHIPMENT.

if(1 == $potem30["remark"]["c4"]){
    $spreadsheet->getActiveSheet()->setCellValue('A37', '6-');
    if(1 == $potem30["remark"]["c5"]){
        $c5 = 'EXCLUDING';
    }elseif (2 == $potem30["remark"]["c5"]){
        $c5 = 'INCLUDING';
    }
    if(1 == $potem30["remark"]["c6"]){
        $c6 = '  TEST CHARGES';
    }elseif (2 == $potem30["remark"]["c6"]){
        $c6 = '  SURCHARGE';
    }
    $spreadsheet->getActiveSheet()->setCellValue('B37', 'PRice '.$c5.$c6);

}

$spreadsheet->getActiveSheet()->setCellValue('B39', 'ANY CONTRARY REPLIED WITHIN '.$potem30["remark"]["c7"].', THIS CONTRACT IS VALID.');

if(1 == $potem30["remark"]["c8"]){
    $c8 = 'EXCLUDING';
}elseif (2 == $potem30["remark"]["c8"]){
    $c8 = 'INCLUDING';
}
$spreadsheet->getActiveSheet()->setCellValue('B40', $c8.'VAT INVOICE');
$spreadsheet->getActiveSheet()->setCellValue('B41', 'ORDER  NO（'.$potem30["remark"]["c9"].')');

//$toaddr = array('Z9','B11','T11','B13','T13','F15','O15','Y15','B23','B34','U23','B41','K41','U41','Z41','AE41','J43','AC43');  //,'C12','D12','E12','F12','G12','H12','I12','J13','K12','K13','J5'
//
//for($i = 1,$y = 0; $i <= count($toaddr) ; $i++ ,$y++){
//
//    $sheet->setCellValue($toaddr[$y],  $potem30["toaddr"]["a".$i]);
//
//}
//
//
////中部form
//
//$nowcol = 14;
////////$spreadsheet->getActiveSheet()->mergeCells("A{$nowcol}:F{$nowcol}");
//$spreadsheet->getActiveSheet()->setCellValue('B'.$nowcol, $potem30["orderform"]["midpono"]);
////////$spreadsheet->getActiveSheet()->setCellValue('I'.$nowcol, $potem30["invoiceform"]["amout"]);
////
////
//for($x = 0 ,$c = 1; $c <= $potem30["orderform"]["formnum"]; $x++ ,$c++){
//
//    $f19 = 15 + 1 * $x;
//
//    $spreadsheet->getActiveSheet()->mergeCells("A{$f19}:B{$f19}");
//    $spreadsheet->getActiveSheet()->mergeCells("C{$f19}:G{$f19}");
//
//
//    $formarr = array('A'.$f19,'C'.$f19);
//
//    for($i = 1,$y = 0; $i <= $potem30["orderform"]["brrnum"] ; $i++ ,$y++){
//
//        $sheet->setCellValue($formarr[$y],  $potem30["orderform"]['b'.$i][$x]);
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
//$nowcol = $potem30["orderform"]["formnum"] > 4 ? ($nowcol + 1) : 21;
//$spreadsheet->getActiveSheet()->setCellValue('C'.$nowcol, $potem30["toaddr"]["a5"]);
//$nowcol++;
//
//$spreadsheet->getActiveSheet()->setCellValue('C'.$nowcol, $potem30["toaddr"]["a6"]);
//$nowcol++;
//$spreadsheet->getActiveSheet()->setCellValue('C'.$nowcol, $potem30["toaddr"]["a7"]);
//$nowcol++;
//$spreadsheet->getActiveSheet()->setCellValue('A'.$nowcol, $potem30["toaddr"]["a8"]);
//$nowcol++;
//$nowcol++;
//$nowcol++;
//$nowcol++;
//$nowcol++;
//$nowcol++;
//$nowcol++;
//
//$spreadsheet->getActiveSheet()->setCellValue('C'.$nowcol, $potem30["remark"]["c1"]);
//$nowcol++;
//$spreadsheet->getActiveSheet()->setCellValue('C'.$nowcol, $potem30["remark"]["c2"]);

////
////$spreadsheet->getActiveSheet()->setCellValue('B'.$nowcol, 'E-MAIL (電郵): '.$potem30["remark"]["c3"]);
////$spreadsheet->getActiveSheet()->setCellValue('I'.$nowcol, $potem30["remark"]["c4"]);
//////$spreadsheet->getActiveSheet()->mergeCells("A{$nowcol}:E{$nowcol}");
//$spreadsheet->getActiveSheet()->setCellValue('O'.$nowcol, $potem30["remark"]["c3"]);
//$nowcol++;
//
//$spreadsheet->getActiveSheet()->setCellValue('O'.$nowcol, $potem30["remark"]["c4"]);

$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

unset($_SESSION['potem30'] ); //注销SESSION

$output=  ($_GET['action'] == 'formdown' )? 1:0;
$nt = date("YmdHis",time()); //转换为日期。
$filenameout = 'potem30out'.$nt.'.xlsx';
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

