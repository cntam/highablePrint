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

$shipp1 =  $_SESSION['shippingp1'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/shipingp1.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
//$spreadsheet->getActiveSheet()->getDefaultColumnDimension()->setWidth(20);  //设置默认列宽
$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(19);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(19);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(19);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(19);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(19);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(19);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('H')->setWidth(19);  //列宽度
$spreadsheet->getDefaultStyle()->getFont()->setSize(7);


$styleArrayl = [

    'borders' => [

        'left' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],

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

//$spreadsheet->getActiveSheet()->getCell('A7')->setValue("hello\nworld");

/* textarea文字模块*/
$spreadsheet->getActiveSheet()->getCell('A7')->setValue($shipp1["foraccount"]);
$spreadsheet->getActiveSheet()->getStyle('A7')->getAlignment()->setWrapText(true);  //在单元格中写入换行符“\ n”（ALT +“Enter”）
/* 文字模块*/

/* textarea文字模块*/
$spreadsheet->getActiveSheet()->getCell('C7')->setValue($shipp1["consignee"]);
$spreadsheet->getActiveSheet()->getStyle('C7')->getAlignment()->setWrapText(true);  //在单元格中写入换行符“\ n”（ALT +“Enter”）
/* 文字模块*/

$sheet->setCellValue('G6',  $shipp1["invoice"]);
$sheet->setCellValue('F8',  $shipp1["shipdate"]);

//头部data
$listarr = array('G9','A11','C12','D12','E12','F11','H11','A14','C14','D14','E14','F14');

for($i= 0; $i <= 11 ; $i++){
    $sheet->setCellValue($listarr[$i],  $shipp1["shipdata"]['a'.$i]);
}





//中部form
$nowcol = 19;

for($x = 0 ,$c = 1; $x <= $shipp1["shipform"]["formnum"] ; $x++ ,$c++){

$f19 = 19 + 4 * $x;
$f20 = 20 + 4 * $x;
$f21 = 21 + 4 * $x;
$f22 = 22 + 4 * $x;
$formarr = array('A'.$f19,'A'.$f20,'B'.$f20,'C'.$f20,'D'.$f20,'F'.$f20,'G'.$f20,'H'.$f20,'A'.$f21,'D'.$f21);
    $spreadsheet->getActiveSheet()->mergeCells("D{$f20}:E{$f20}");
    $spreadsheet->getActiveSheet()->mergeCells("D{$f21}:E{$f21}");
    for($i= 1,$y = 0; $i <= 10 ; $i++ ,$y++){

        $sheet->setCellValue($formarr[$y],  $shipp1["shipform"]['b'.$i][$x]);
    }


    $nowcol = 19 +   4 * $c;

    $spreadsheet->getActiveSheet()->getCell('A1')->setValue($nowcol);


}

$spreadsheet->getActiveSheet()->getCell('H'.$nowcol)->setValue($shipp1["shipform"]["fortotal"]);
$nowcol++;
$nowcol++;

$spreadsheet->getActiveSheet()->mergeCells("B{$nowcol}:F{$nowcol}");
$spreadsheet->getActiveSheet()->getCell('B'.$nowcol)->setValue($shipp1["shipform"]["formremark"]);
$nowcol++;
$nowcol++;
$nowcol++;

$spreadsheet->getActiveSheet()->getCell('A'.$nowcol)->setValue('MID CODE：');
$spreadsheet->getActiveSheet()->mergeCells("B{$nowcol}:F{$nowcol}");
$spreadsheet->getActiveSheet()->getCell('B'.$nowcol)->setValue($shipp1["shipbottom"]["c1"]);
$spreadsheet->getActiveSheet()->getStyle("B{$nowcol}:F{$nowcol}")->applyFromArray($styleArraybu);
$nowcol++;

$spreadsheet->getActiveSheet()->getCell('A'.$nowcol)->setValue('COUNTRY OF ORIGIN：');
$spreadsheet->getActiveSheet()->mergeCells("B{$nowcol}:F{$nowcol}");
$spreadsheet->getActiveSheet()->getCell('B'.$nowcol)->setValue($shipp1["shipbottom"]["c2"]);
$spreadsheet->getActiveSheet()->getStyle("B{$nowcol}:F{$nowcol}")->applyFromArray($styleArraybu);
$nowcol++;
$nowcol++;





//边栏样式
$spreadsheet->getActiveSheet()->getStyle("A19:A{$nowcol}")->applyFromArray($styleArrayl);
$spreadsheet->getActiveSheet()->getStyle("H19:H{$nowcol}")->applyFromArray($styleArrayr);
$spreadsheet->getActiveSheet()->getStyle("A{$nowcol}:H{$nowcol}")->applyFromArray($styleArraybu);





$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

//unset($_SESSION['shipp1'] ); //注销SESSION

$output=  ($_GET['action'] == 'formdown' )? 1:0;
$nt = date("YmdHis",time()); //转换为日期。
$filenameout = 'shipp1out'.$nt.'.xlsx';
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

