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

$intem1 =  $_SESSION['invoiceTem1'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/invoiceTem1.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft YaHei');
//$spreadsheet->getActiveSheet()->getDefaultColumnDimension()->setWidth(20);  //设置默认列宽
$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(9);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(9);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(11);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(13);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(14);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(14);  //列宽度
$spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(28);  //列宽度
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
$spreadsheet->getActiveSheet()->setCellValue('F8', $intem1["invoicedata"]["invoiceNumber"]);
$spreadsheet->getActiveSheet()->setCellValue('F9', $intem1["invoicedata"]["ups"]);
$spreadsheet->getActiveSheet()->setCellValue('I11', $intem1["invoicedate"]);
$spreadsheet->getActiveSheet()->setCellValue('B11', $intem1["tosb"]);

////$spreadsheet->getActiveSheet()->getCell('A7')->setValue("hello\nworld");
//
///* textarea文字模块*/
//$spreadsheet->getActiveSheet()->getCell('A7')->setValue($intem1["foraccount"]);
//$spreadsheet->getActiveSheet()->getStyle('A7')->getAlignment()->setWrapText(true);  //在单元格中写入换行符“\ n”（ALT +“Enter”）
///* 文字模块*/
//
///* textarea文字模块*/
//$spreadsheet->getActiveSheet()->getCell('C7')->setValue($intem1["consignee"]);
//$spreadsheet->getActiveSheet()->getStyle('C7')->getAlignment()->setWrapText(true);  //在单元格中写入换行符“\ n”（ALT +“Enter”）
///* 文字模块*/



//头部data
$listarr = array('B12','B13','B14','B15','B16','D17','D18','B19');

for($i= 0,$l = 1; $i < 8 ; $i++,$l++){
    $sheet->setCellValue($listarr[$i],  $intem1["invoicedata"]['a'.$l]);
}





//中部form
$nowcol = 21;
$spreadsheet->getActiveSheet()->setCellValue('H'.$nowcol, $intem1["invoiceform"]["price"]);
$spreadsheet->getActiveSheet()->setCellValue('I'.$nowcol, $intem1["invoiceform"]["amout"]);

$nowcol++;

for($x = 0 ,$c = 1; $c <= $intem1["invoiceform"]["formnum"]; $x++ ,$c++){

$f19 = 22 + 1 * $x;

$formarr = array('A'.$f19,'B'.$f19,'C'.$f19,'D'.$f19,'E'.$f19,'F'.$f19,'G'.$f19,'H'.$f19,'I'.$f19);

    for($i = 1,$y = 0; $i <= $intem1["invoiceform"]["brrnum"] ; $i++ ,$y++){

        $sheet->setCellValue($formarr[$y],  $intem1["invoiceform"]['b'.$i][$x]);

    }


    $nowcol = 22 +   1 * $c;



    if($x >12){
        $spreadsheet->getActiveSheet()->insertNewRowBefore($nowcol, 1);
    }

}
$nowcol = $intem1["invoiceform"]["formnum"] > 12 ? ($nowcol + 1) : 36;
//$spreadsheet->getActiveSheet()->getCell('A1')->setValue($nowcol);

$spreadsheet->getActiveSheet()->setCellValue('A'.$nowcol, $intem1["invoiceform"]["coltb"]);
$spreadsheet->getActiveSheet()->setCellValue('I'.$nowcol, $intem1["invoiceform"]["coltc"]);

$nowcol++;
$nowcol++;
$spreadsheet->getActiveSheet()->getCell('D'.$nowcol)->setValue($intem1["invoiceform"]["formremark"]);
$spreadsheet->getActiveSheet()->getStyle('D'.$nowcol)->applyFromArray($styleArray1);

$nowcol++;
$nowcol++;
$nowcol++;
$spreadsheet->getActiveSheet()->mergeCells("E{$nowcol}:H{$nowcol}");
$spreadsheet->getActiveSheet()->getCell('E'.$nowcol)->setValue($intem1["remark"]["bottomremark"]);
$spreadsheet->getActiveSheet()->getStyle("F{$nowcol}:H{$nowcol}")->applyFromArray($styleArray1);
$nowcol++;
$nowcol++;


for($b = 1 ; $b<= $intem1["remark"]["crrnum"] ; $b++ ){
    $spreadsheet->getActiveSheet()->mergeCells("F{$nowcol}:H{$nowcol}");
    $spreadsheet->getActiveSheet()->getCell('F'.$nowcol)->setValue($intem1["remark"]["c".$b]);
    $spreadsheet->getActiveSheet()->getStyle("F{$nowcol}:H{$nowcol}")->applyFromArray($styleArray1);
    $nowcol++;


}



////边栏样式
//$spreadsheet->getActiveSheet()->getStyle("A19:A{$nowcol}")->applyFromArray($styleArrayl);
//$spreadsheet->getActiveSheet()->getStyle("H19:H{$nowcol}")->applyFromArray($styleArrayr);
//$spreadsheet->getActiveSheet()->getStyle("A{$nowcol}:H{$nowcol}")->applyFromArray($styleArraybu);





$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

//unset($_SESSION['shipp1'] ); //注销SESSION

$output=  ($_GET['action'] == 'formdown' )? 1:0;
$nt = date("YmdHis",time()); //转换为日期。
$filenameout = 'intem1out'.$nt.'.xlsx';
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

