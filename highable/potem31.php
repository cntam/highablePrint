<?php

require_once 'aidenfunc.php';

$potem31 =  $_SESSION['potem31'];
$pop1 =  $_SESSION['potem31'];


//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/potem30.xlsx');

$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle("sheet1");
$spreadsheet->getDefaultStyle()->getFont()->setName('SimSun');
//$sheet->getDefaultColumnDimension()->setWidth(30);  //设置默认列宽
for($j=0;$j<=8;$j++){
    $col = chr(65 + $j);
    $sheet->getColumnDimension($col)->setWidth(15);  //列宽度
}

//$sheet->getColumnDimension('A')->setWidth(25);  //列宽度
//$sheet->getColumnDimension('B')->setWidth(5);  //列宽度
//$sheet->getColumnDimension('C')->setWidth(15);  //列宽度

$spreadsheet->getDefaultStyle()->getFont()->setSize(7);

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
setCell($sheet,"A1",$potem31["remark"]['poheader']['poheada1'],$noborderCenter);
setCell($sheet,"A2",$potem31["remark"]['poheader']['poheada2'],$noborderCenter);
setCell($sheet,"A3",$potem31["remark"]['poheader']['poheada3'],$noborderCenter);
setCell($sheet,"A4",$potem31["remark"]['poheader']['poheada4'].$potem31["remark"]['poheader']['poheada5'],$noborderCenter);

$sheet->setCellValue('A9', $potem31["tosb"]);
$sheet->setCellValue('B18', $potem31["podate"]);

//if(1 == $potem31["toaddr"]["a1"]){
//  $titlecon = 'HIGH ABLE INVESTMENT LIMITED';
//}elseif (2 == $potem31["toaddr"]["a1"]){
//    $titlecon = 'IRONDALE FASHION INTERNATIONAL LIMITED';
//}
//$sheet->setCellValue('A1', $titlecon);
$sheet->setCellValue('G9', $potem31["toaddr"]["a2"]);
$sheet->setCellValue('A10', $potem31["toaddr"]["a3"]);
$sheet->setCellValue('A11', $potem31["toaddr"]["a4"]);
$sheet->setCellValue('A12', $potem31["toaddr"]["a5"]);
$sheet->setCellValue('A13', '电话：'.$potem31["toaddr"]["a6"]); //电话：0571-86312008 传真：0571-86312007
$sheet->setCellValue('A14', '传真：'.$potem31["toaddr"]["a7"]); //电话：0571-86312008 传真：0571-86312007
$sheet->setCellValue('A15', $potem31["toaddr"]["a8"]);

if(1 == $potem31["toaddr"]["a9"]){
    $amount = 'Amount(RMB)';
}elseif (2 == $potem31["toaddr"]["a9"]){
    $amount = 'Amount(HKD)';
}elseif (3 == $potem31["toaddr"]["a9"]){
    $amount = 'Amount(USD)';
}
$sheet->setCellValue('I19', $amount);
$sheet->setCellValue('A21', $potem31["orderform"]['b1'][0]);
$sheet->setCellValue('G21', $potem31["orderform"]['b2'][0]);
$sheet->setCellValue('I21', $potem31["orderform"]['b3'][0]);
$sheet->setCellValue('B22', $potem31["orderform"]['b4'][0]);
$sheet->setCellValue('G22', $potem31["orderform"]['b5'][0]);
$sheet->setCellValue('B23', $potem31["orderform"]['b6'][0]);

if(1 == $potem31["toaddr"]["a11"]){
    $um = 'U/M';
}elseif (2 == $potem31["toaddr"]["a11"]){
    $um = 'U/Y';
}
$sheet->setCellValue('H25', $um);

$sheet->setCellValue('G27', $potem31["toaddr"]["a16"]);
$sheet->setCellValue('H27', $potem31["toaddr"]["a17"]);

$sheet->setCellValue('A28', 'Total   Amount  ：'.$potem31["toaddr"]["a18"]);
$sheet->setCellValue('C28', $potem31["toaddr"]["a19"]);
$sheet->setCellValue('A29', 'Payment  Terms：'.$potem31["toaddr"]["a20"]);
$sheet->setCellValue('A30', 'Price   Terms    ：'.$potem31["toaddr"]["a21"]);

/**
 * remark
 */
$sheet->setCellValue('A32', $pop1["remark"]["c2"][0]);
$sheet->setCellValue('B32', 'AMOUNT & QUANTITY WITHIN THE TOLERANCE OF '.$potem31["remark"]["c2"][1].' MORE OR LESS IS ONLY ALLOWED.');

$sheet->setCellValue('A33', $pop1["remark"]["c3"][0]);
$sheet->setCellValue('B33', 'YOUR PARTY MUST TAKE FULL RESPONSIBILITY FOR ANY DELAY OF SHIPMENT.');

$sheet->setCellValue('A34', $pop1["remark"]["c4"][0]);
$sheet->setCellValue('B34', 'AZO  FREE');

$sheet->setCellValue('A35', $pop1["remark"]["c5"][0]);
$sheet->setCellValue('B35', 'ALL  PERFORMANCES SHOULD MEET OUR REQUIREMENTS（AS PER ATTACHED）.');

$sheet->setCellValue('A36', $pop1["remark"]["c6"][0]);
$sheet->setCellValue('B36', 'YOU HAVE TO SUBMIT '.$pop1["remark"]["c6"][1].'  SHIPMENT SAMPLE FOR OUR APPROVAL BEFORE '.$pop1["remark"]["c6"][2].' OF SHIPMENT.');


if(1 == $potem31["remark"]["c7"][1]){
    $sheet->setCellValue('A37', $pop1["remark"]["c7"][0]);
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
    $sheet->setCellValue('B37', 'Price '.$c5.$c6);
}

//$sheet->setCellValue('B39', 'ANY CONTRARY REPLIED WITHIN '.$potem31["remark"]["c7"].', THIS CONTRACT IS VALID.');

$sheet->setCellValue('B40', 'PLEASE  CONFIRM  AND  COUNTER-SIGN  BY  RETURN.OTHERWISE,IF WE DO NOT RECEIVE ANY CONTRARY REPLIED WITHIN '.$potem31["remark"]["c11"].',THIS CONTRACT IS VALID.');
$sheet->getStyle('B40')->applyFromArray($noborderLeft);

if(1 == $potem31["remark"]["c12"]){
    $c8 = 'EXCLUDING';
}elseif (2 == $potem31["remark"]["c12"]){
    $c8 = 'INCLUDING';
}
$sheet->setCellValue('B42', $c8.'  VAT INVOICE');

//$sheet->setCellValue('B42', $potem31["remark"]["c10"].'VAT INVOICE');
$sheet->setCellValue('B43', 'ORDER  NO: '.$potem31["remark"]["c13"]);

$row = 44;
if(is_array($potem31["remark"]["c14"])){
    if(count($potem31["remark"]["c14"]) > 1){
        foreach ($potem31["remark"]["c14"] as $item=>$value){

            if($item >1){
                $sheet->insertNewRowBefore($row, 1);
            }

            $sheet->setCellValue('A'. $row, $value );
            $sheet->getStyle('A'. $row)->applyFromArray($styleArray3);
            //$sheet->getStyle('A'. $row)->applyFromArray($noborderLeft);
            //$sheet->getStyle('A'. $row)->getAlignment()->setWrapText(true);

            $sheet->mergeCells("B{$row}:H{$row}");
            $sheet->setCellValue('B'. $row, $potem31["remark"]["c15"][$item]);
            $sheet->getStyle('B'. $row)->applyFromArray($noborderLeft);
            $sheet->getStyle('B'. $row)->getAlignment()->setWrapText(true);

            $row++;
        }
    }
}


/**
 *   remark中间增加行
 */
$row = 38;
if(is_array($pop1["remark"]["c8"])){
    if(count($pop1["remark"]["c8"]) > 1){
        foreach ($pop1["remark"]["c8"] as $item=>$value){

            if($item >0){
                $sheet->insertNewRowBefore($row, 1);
            }

            $sheet->setCellValue('A'. $row, $value );
            $sheet->getStyle('A'. $row)->applyFromArray($styleArray3);
            //$sheet->getStyle('A'. $row)->applyFromArray($noborderLeft);


            $sheet->mergeCells("B{$row}:H{$row}");
            $sheet->setCellValue('B'. $row, $pop1["remark"]["c9"][$item]);
            $sheet->getStyle('B'. $row)->applyFromArray($noborderLeft);
            $sheet->getStyle('B'. $row)->getAlignment()->setWrapText(true);

            $row++;
        }


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
            $sheet->insertNewRowBefore($row, 1);
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

        $sheet->insertNewRowBefore($row, 1);
        $sheet->mergeCells("B{$row}:F{$row}");
        $sheet->setCellValue('B'. $row, $value);
        $row++;
    }
}

/**
 * PO详情 附加
 */

$sheet->getPageSetup()
    ->setOrientation(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_PORTRAIT);  //竖放置
$sheet->getPageSetup()
    ->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);  //A4
$sheet->getPageSetup()->setFitToPage(true); //将工作表调整为一页

unset($_SESSION['potem31'] ); //注销SESSION



$filenameout = 'PO_'.$potem31['pono'];
outExcel($spreadsheet,$filenameout);

