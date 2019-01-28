<?php
session_start();
header("Content-type: text/html; charset=utf-8");
/*港源行國際有限公司*/

require_once('autoloadconfig.php');  //判断是否在线

require_once ('img.php');

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;

$packinglistTem1 =  $_SESSION['packinglist'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/packinglistTem1.xlsx');
$sheet = $spreadsheet->getActiveSheet();
//样式
$styleArray1 = [
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
        'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
    ],
];

//填数据
//头部
$row = 3;
if ($packinglistTem1["invoicedata"]["acolnum"] > 0) {
    for ($x = 1; $x <= 6; $x++) {
        $sheet->setCellValue('G'.$row, $packinglistTem1["invoicedata"]['a'.$x]);
        $row++;
    }
    $row = 3;
    for ($x = 7; $x <= 12; $x++) {
        $sheet->setCellValue('Z'.$row, $packinglistTem1["invoicedata"]['a'.$x]);
        $row++;
    }
}
//中间表格
$sheet->setCellValue('E11', $packinglistTem1["invoicedata"]["a13"]);
$sheet->setCellValue('S11', $packinglistTem1["invoicedata"]["a14"]);
$sheet->setCellValue('X11', $packinglistTem1["invoicedata"]["a15"]);

$sheet->setCellValue('AG8', $packinglistTem1["invoicedata"]["a16"]);
$sheet->setCellValue('AG9', 'G.W.: '.$packinglistTem1["invoicedata"]["a17"].'KGS');
$sheet->setCellValue('AG10', 'N.W.:  '.$packinglistTem1["invoicedata"]["a18"].'KGS');

//size格
{
    $row = 11;
    if ($packinglistTem1["invoiceform"]["brownum"] > 0) {
        for ($a = 0; $a < $packinglistTem1["invoiceform"]["brownum"]; $a++) {
            $sizeContent = $packinglistTem1["invoiceform"]["b1"][$a].'X'.$packinglistTem1["invoiceform"]["b2"][$a].'X'.$packinglistTem1["invoiceform"]["b3"][$a].'CM';
            $sheet->setCellValue('AF'.$row, 'SIZE:'.$sizeContent);
            $mergeAF = 'AF'.$row;
            $mergeAK = 'AK'.$row;
            $sheet->mergeCells("$mergeAF:$mergeAK");
            $sheet->getStyle($mergeAF)->applyFromArray($styleArray1);
            $row++;
        }
//        $row++;
        $sheet->setCellValue('AF'.$row, 'CMB:  '.$packinglistTem1["invoiceform"]["b4"].'m³');
    }
}

//Size Breakdown
$sheet->setCellValue('R13', $packinglistTem1["remark"]["clist"]["c1"][0]);
$sheet->setCellValue('T13', $packinglistTem1["remark"]["clist"]["c1"][1]);
$sheet->setCellValue('U13', $packinglistTem1["remark"]["clist"]["c1"][2]);
$sheet->setCellValue('V13', $packinglistTem1["remark"]["clist"]["c1"][3]);
$sheet->setCellValue('W13', $packinglistTem1["remark"]["clist"]["c1"][4]);
$sheet->setCellValue('X13', $packinglistTem1["remark"]["clist"]["c1"][5]);
$sheet->setCellValue('Y13', $packinglistTem1["remark"]["clist"]["c1"][6]);

$sheet->setCellValue('T25', $packinglistTem1["remark"]["dlist"]["d8"]);
$sheet->setCellValue('V25', $packinglistTem1["remark"]["dlist"]["d9"]);

//动态

if ($packinglistTem1["remark"]["dlist"]["dnum"] > 0) {
    $arr = array('C', 'J', 'M', 'P', 'T', 'V');

    for ($a = 0, $b = 2; $a < count($arr); $a++, $b++) {
        $row = 19;
        foreach ($packinglistTem1["remark"]["dlist"]['d'.$b] as $item=>$value) {
            if (($item > 5)&&($b == 2)) {
                $sheet->insertNewRowBefore($row, 1);
//                新增行的样式（合并单元格）
                $contextArr = array('C'.$row,'I'.$row, 'J'.$row, 'L'.$row, 'M'.$row, 'O'.$row, 'P'.$row, 'S'.$row,
                    'T'.$row, 'U'.$row, 'V'.$row, 'W'.$row, 'Y'.$row, 'Z'.$row, 'AA'.$row, 'AC'.$row, 'AD'.$row, 'AF'.$row,);
                $sheet->mergeCells("$contextArr[0]:$contextArr[1]");
                $sheet->mergeCells("$contextArr[2]:$contextArr[3]");
                $sheet->mergeCells("$contextArr[4]:$contextArr[5]");
                $sheet->mergeCells("$contextArr[6]:$contextArr[7]");
                $sheet->mergeCells("$contextArr[8]:$contextArr[9]");
                $sheet->mergeCells("$contextArr[10]:$contextArr[11]");
                $sheet->mergeCells("$contextArr[12]:$contextArr[13]");
                $sheet->mergeCells("$contextArr[14]:$contextArr[15]");
                $sheet->mergeCells("$contextArr[16]:$contextArr[17]");

            }
            $sheet->setCellValue($arr[$a].$row, $value);
            $row++;
        }
    }
}
if ($packinglistTem1["remark"]["elist"]["enum"] > 0) {
    $arr = array('Y', 'AA', 'AD', 'AG', 'AH');

    for ($a = 0, $b = 1; $a < count($arr); $a++, $b++) {
        $row = 19;
        foreach ($packinglistTem1["remark"]["elist"]['e'.$b] as $item=>$value) {
            $sheet->setCellValue($arr[$a].$row, $value);
            $row++;
        }
    }
}

if ($packinglistTem1["remark"]["clist"]["cnum"] > 0) {
    $arr = array('A', 'C', 'E', 'K', 'N', 'R', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AB', 'AD');

    for ($a = 0, $b = 2; $a < count($arr); $a++, $b++) {
        $row = 14;
        foreach ($packinglistTem1["remark"]["clist"]['c'.$b] as $item=>$value) {
            if (($item > 0)&&($b == 2)) {
                $sheet->insertNewRowBefore($row, 1);
                $contextArr = array('A'.$row,'B'.$row, 'C'.$row, 'D'.$row, 'E'.$row, 'J'.$row, 'K'.$row, 'M'.$row,
                    'N'.$row, 'Q'.$row, 'R'.$row, 'S'.$row, 'Z'.$row, 'AA'.$row, 'AB'.$row, 'AC'.$row, 'AD'.$row, 'AE'.$row,);
                $sheet->mergeCells("$contextArr[0]:$contextArr[1]");
                $sheet->mergeCells("$contextArr[2]:$contextArr[3]");
                $sheet->mergeCells("$contextArr[4]:$contextArr[5]");
                $sheet->mergeCells("$contextArr[6]:$contextArr[7]");
                $sheet->mergeCells("$contextArr[8]:$contextArr[9]");
                $sheet->mergeCells("$contextArr[10]:$contextArr[11]");
                $sheet->mergeCells("$contextArr[12]:$contextArr[13]");
                $sheet->mergeCells("$contextArr[14]:$contextArr[15]");
                $sheet->mergeCells("$contextArr[16]:$contextArr[17]");

            }
            $sheet->setCellValue($arr[$a].$row, $value);
            $row++;
        }
    }

}

$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

//unset($_SESSION['packinglist'] ); //注销SESSION

$output=  ($_GET['action'] == 'formdown' )? 1:0;
$nt = date("md",time()); //转换为日期。
$filenameout = 'Packinglist_KM_'.$nt.'.xlsx';
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
