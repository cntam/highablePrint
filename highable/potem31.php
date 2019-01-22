<?php
session_start();
require_once('autoloadconfig.php');  //判断是否在线

require_once ('img.php');

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Helper\Html as HtmlHelper; // html 解析器

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;

$potem31 =  $_SESSION['potem31'];
$pop1 =  $_SESSION['potem31'];


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

$styleArray3 = [
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
        'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
        'wrapText' => true,
        'ShrinkToFit'=>true,
    ],
    'font' => [
        'Size' => '5',
    ]

];

//填数据
$spreadsheet->getActiveSheet()->setCellValue('A9', $potem31["tosb"]);
$spreadsheet->getActiveSheet()->setCellValue('B18', $potem31["podate"]);

if(1 == $potem31["toaddr"]["a1"]){
  $titlecon = 'HIGH ABLE INVESTMENT LIMITED';
}elseif (2 == $potem31["toaddr"]["a1"]){
    $titlecon = 'IRONDALE FASHION INTERNATIONAL LIMITED';
}
$spreadsheet->getActiveSheet()->setCellValue('A1', $titlecon);
$spreadsheet->getActiveSheet()->setCellValue('G9', $potem31["toaddr"]["a2"]);
$spreadsheet->getActiveSheet()->setCellValue('A10', $potem31["toaddr"]["a3"]);
$spreadsheet->getActiveSheet()->setCellValue('A11', $potem31["toaddr"]["a4"]);
$spreadsheet->getActiveSheet()->setCellValue('A12', $potem31["toaddr"]["a5"]);
$spreadsheet->getActiveSheet()->setCellValue('A13', '电话：'.$potem31["toaddr"]["a6"]); //电话：0571-86312008 传真：0571-86312007
$spreadsheet->getActiveSheet()->setCellValue('A14', '传真：'.$potem31["toaddr"]["a7"]); //电话：0571-86312008 传真：0571-86312007
$spreadsheet->getActiveSheet()->setCellValue('A15', $potem31["toaddr"]["a8"]);

if(1 == $potem31["toaddr"]["a9"]){
    $amount = 'Amount(RMB)';
}elseif (2 == $potem31["toaddr"]["a9"]){
    $amount = 'Amount(HKD)';
}elseif (3 == $potem31["toaddr"]["a9"]){
    $amount = 'Amount(USD)';
}
$spreadsheet->getActiveSheet()->setCellValue('I19', $amount);
$spreadsheet->getActiveSheet()->setCellValue('A21', $potem31["orderform"]['b1'][0]);
$spreadsheet->getActiveSheet()->setCellValue('G21', $potem31["orderform"]['b2'][0]);
$spreadsheet->getActiveSheet()->setCellValue('I21', $potem31["orderform"]['b3'][0]);
$spreadsheet->getActiveSheet()->setCellValue('B22', $potem31["orderform"]['b4'][0]);
$spreadsheet->getActiveSheet()->setCellValue('G22', $potem31["orderform"]['b5'][0]);
$spreadsheet->getActiveSheet()->setCellValue('B23', $potem31["orderform"]['b6'][0]);

if(1 == $potem31["toaddr"]["a11"]){
    $um = 'U/M';
}elseif (2 == $potem31["toaddr"]["a11"]){
    $um = 'U/Y';
}
$spreadsheet->getActiveSheet()->setCellValue('H25', $um);

$spreadsheet->getActiveSheet()->setCellValue('G27', $potem31["toaddr"]["a16"]);
$spreadsheet->getActiveSheet()->setCellValue('H27', $potem31["toaddr"]["a17"]);

$spreadsheet->getActiveSheet()->setCellValue('A28', 'Total   Amount  ：'.$potem31["toaddr"]["a18"]);
$spreadsheet->getActiveSheet()->setCellValue('C28', $potem31["toaddr"]["a19"]);
$spreadsheet->getActiveSheet()->setCellValue('A29', 'Payment  Terms：'.$potem31["toaddr"]["a20"]);
$spreadsheet->getActiveSheet()->setCellValue('A30', 'Price   Terms    ：'.$potem31["toaddr"]["a21"]);

/**
 * remark
 */
$spreadsheet->getActiveSheet()->setCellValue('A32', $pop1["remark"]["c2"][0]);
$spreadsheet->getActiveSheet()->setCellValue('B32', 'AMOUNT & QUANTITY WITHIN THE TOLERANCE OF '.$potem31["remark"]["c2"][1].' MORE OR LESS IS ONLY ALLOWED.');

$spreadsheet->getActiveSheet()->setCellValue('A33', $pop1["remark"]["c3"][0]);
$spreadsheet->getActiveSheet()->setCellValue('B33', 'YOUR PARTY MUST TAKE FULL RESPONSIBILITY FOR ANY DELAY OF SHIPMENT.');

$spreadsheet->getActiveSheet()->setCellValue('A34', $pop1["remark"]["c4"][0]);
$spreadsheet->getActiveSheet()->setCellValue('B34', 'AZO  FREE');

$spreadsheet->getActiveSheet()->setCellValue('A35', $pop1["remark"]["c5"][0]);
$spreadsheet->getActiveSheet()->setCellValue('B35', 'ALL  PERFORMANCES SHOULD MEET OUR REQUIREMENTS（AS PER ATTACHED）.');

$spreadsheet->getActiveSheet()->setCellValue('A36', $pop1["remark"]["c6"][0]);
$spreadsheet->getActiveSheet()->setCellValue('B36', 'YOU HAVE TO SUBMIT '.$pop1["remark"]["c6"][1].'  SHIPMENT SAMPLE FOR OUR APPROVAL BEFORE '.$pop1["remark"]["c6"][2].' OF SHIPMENT.');


if(1 == $potem31["remark"]["c7"][1]){
    $spreadsheet->getActiveSheet()->setCellValue('A37', $pop1["remark"]["c7"][0]);
    if(1 == $pop1["remark"]["c7"][2]){
        $c5 = 'EXCLUDING';
    }elseif (2 == $pop1["remark"]["c7"][2]){
        $c5 = 'INCLUDING';
    }
    if(1 == $pop1["remark"]["c7"][3]){
        $c6 = '  TEST CHARGES';
    }elseif (2 == $pop1["remark"]["c7"][3]){
        $c6 = '  SURCHARGE';
    }
    $spreadsheet->getActiveSheet()->setCellValue('B37', 'Price '.$c5.$c6);
}

//$spreadsheet->getActiveSheet()->setCellValue('B39', 'ANY CONTRARY REPLIED WITHIN '.$potem31["remark"]["c7"].', THIS CONTRACT IS VALID.');

$spreadsheet->getActiveSheet()->setCellValue('B40', 'PLEASE  CONFIRM  AND  COUNTER-SIGN  BY  RETURN.OTHERWISE,IF WE DO NOT RECEIVE ANY CONTRARY REPLIED WITHIN '.$potem31["remark"]["c11"].',THIS CONTRACT IS VALID.');
$spreadsheet->getActiveSheet()->getStyle('B40')->applyFromArray($styleArray2);

if(1 == $potem31["remark"]["c12"]){
    $c8 = 'EXCLUDING';
}elseif (2 == $potem31["remark"]["c12"]){
    $c8 = 'INCLUDING';
}
$spreadsheet->getActiveSheet()->setCellValue('B42', $c8.'  VAT INVOICE');

//$spreadsheet->getActiveSheet()->setCellValue('B42', $potem31["remark"]["c10"].'VAT INVOICE');
$spreadsheet->getActiveSheet()->setCellValue('B43', 'ORDER  NO: '.$potem31["remark"]["c13"]);

$row = 44;
if(count($potem31["remark"]["c14"]) > 1){
    foreach ($potem31["remark"]["c14"] as $item=>$value){

        if($item >1){
            $spreadsheet->getActiveSheet()->insertNewRowBefore($row, 1);
        }

        $sheet->setCellValue('A'. $row, $value );
        $spreadsheet->getActiveSheet()->getStyle('A'. $row)->applyFromArray($styleArray3);
        //$spreadsheet->getActiveSheet()->getStyle('A'. $row)->applyFromArray($styleArray2);
        //$spreadsheet->getActiveSheet()->getStyle('A'. $row)->getAlignment()->setWrapText(true);

        $spreadsheet->getActiveSheet()->mergeCells("B{$row}:H{$row}");
        $sheet->setCellValue('B'. $row, $potem31["remark"]["c15"][$item]);
        $spreadsheet->getActiveSheet()->getStyle('B'. $row)->applyFromArray($styleArray2);
        $spreadsheet->getActiveSheet()->getStyle('B'. $row)->getAlignment()->setWrapText(true);

        $row++;
    }
}

/**
 *   remark中间增加行
 */
$row = 38;
if(count($pop1["remark"]["c8"]) > 1){
    foreach ($pop1["remark"]["c8"] as $item=>$value){

        if($item >0){
            $spreadsheet->getActiveSheet()->insertNewRowBefore($row, 1);
        }

        $sheet->setCellValue('A'. $row, $value );
        $spreadsheet->getActiveSheet()->getStyle('A'. $row)->applyFromArray($styleArray3);
        //$spreadsheet->getActiveSheet()->getStyle('A'. $row)->applyFromArray($styleArray2);


        $spreadsheet->getActiveSheet()->mergeCells("B{$row}:H{$row}");
        $sheet->setCellValue('B'. $row, $pop1["remark"]["c9"][$item]);
        $spreadsheet->getActiveSheet()->getStyle('B'. $row)->applyFromArray($styleArray2);
        $spreadsheet->getActiveSheet()->getStyle('B'. $row)->getAlignment()->setWrapText(true);

        $row++;
    }


}
/**
 *   remark中间增加行
 */

/**
 * remark
 */


/**
 * 中间报价表格
 */
$row = 26;
if(count($pop1["toaddr"]["a12"]) > 0){
    foreach ($pop1["toaddr"]["a12"] as $item=>$value){

        if($item >0){
            $spreadsheet->getActiveSheet()->insertNewRowBefore($row, 1);
        }

        $sheet->setCellValue('B'. $row, $value);
        $sheet->setCellValue('F'. $row, $pop1["toaddr"]["a13"][$item]);
        $sheet->setCellValue('G'. $row, $pop1["toaddr"]["a14"][$item]);
        $sheet->setCellValue('H'. $row, $pop1["toaddr"]["a15"][$item]);

        $row++;
    }
}

/**
 * 中间报价表格
 */

/**
 * PO详情 附加
 */
$row = 24;
if($pop1["orderform"]["elist"]['elistrow'] > 1){
    foreach ($pop1["orderform"]["elist"]['e1'] as $item=>$value){

        $spreadsheet->getActiveSheet()->insertNewRowBefore($row, 1);
        $spreadsheet->getActiveSheet()->mergeCells("B{$row}:F{$row}");
        $sheet->setCellValue('B'. $row, $value);
        $row++;
    }
}

/**
 * PO详情 附加
 */





$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

unset($_SESSION['potem31'] ); //注销SESSION

require_once 'aidenfunc.php';

$filenameout = 'PO_'.$pop1['shortName'];
outExcel($spreadsheet,$filenameout);

