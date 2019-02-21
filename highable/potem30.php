<?php

require_once 'aidenfunc.php';

$potem30 =  $_SESSION['potem30'];

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


//填数据
//header
setCell($sheet, "A1", $potem30["remark"]["poheader"]["poheada1"], $noborderCenter);
setCell($sheet, "A2", 'Address:'.$potem30["remark"]["poheader"]["poheada2"], $noborderCenter);
setCell($sheet, "A3", $potem30["remark"]["poheader"]["poheada3"], $noborderCenter);
setCell($sheet,"A4",'Tel:'.$potem30["remark"]['poheader']['poheada5'].'Fax:'.$potem30["remark"]['poheader']['poheada5'],$noborderCenter);

$sheet->setCellValue('A9', $potem30["tosb"]);
$sheet->setCellValue('B18', $potem30["podate"]);

//if(1 == $potem30["toaddr"]["a1"]){
//  $titlecon = 'HIGH ABLE INVESTMENT LIMITED';
//}elseif (2 == $potem30["toaddr"]["a1"]){
//    $titlecon = 'IRONDALE FASHION INTERNATIONAL LIMITED';
//}
//$sheet->setCellValue('A1', $titlecon);
setCell($sheet,"A1",$potem30["remark"]['poheader']['poheada1'],$noborderCenter);
setCell($sheet,"A2",$potem30["remark"]['poheader']['poheada2'],$noborderCenter);
setCell($sheet,"A3",$potem30["remark"]['poheader']['poheada3'],$noborderCenter);
setCell($sheet,"A4",$potem30["remark"]['poheader']['poheada4'].$potem30["remark"]['poheader']['poheada5'],$noborderCenter);
//setCell($sheet,"A5",'Attn :'.$potem1["toaddr"]["a9"],$noborderCenter);



$sheet->setCellValue('G9', $potem30["toaddr"]["a2"]);
$sheet->setCellValue('A10', $potem30["toaddr"]["a3"]);
$sheet->setCellValue('A11', $potem30["toaddr"]["a4"]);
$sheet->setCellValue('A12', $potem30["toaddr"]["a5"]);
$sheet->setCellValue('A13', '电话：'.$potem30["toaddr"]["a6"]); //电话：0571-86312008 传真：0571-86312007
$sheet->setCellValue('A14', '传真：'.$potem30["toaddr"]["a7"]); //电话：0571-86312008 传真：0571-86312007
$sheet->setCellValue('A15', $potem30["toaddr"]["a8"]);

if(1 == $potem30["toaddr"]["a9"]){
    $amount = 'Amount(RMB)';
}elseif (2 == $potem30["toaddr"]["a9"]){
    $amount = 'Amount(HKD)';
}elseif (3 == $potem30["toaddr"]["a9"]){
    $amount = 'Amount(USD)';
}
$sheet->setCellValue('I19', $amount);
$sheet->setCellValue('A21', $potem30["orderform"]['b1'][0]);
$sheet->setCellValue('G21', $potem30["orderform"]['b2'][0]);
$sheet->setCellValue('I21', $potem30["orderform"]['b3'][0]);
$sheet->setCellValue('B22', $potem30["orderform"]['b4'][0]);
$sheet->setCellValue('G22', $potem30["orderform"]['b5'][0]);
$sheet->setCellValue('B23', $potem30["orderform"]['b6'][0]);

if(1 == $potem30["toaddr"]["a11"]){
    $um = 'U/M';
}elseif (2 == $potem30["toaddr"]["a11"]){
    $um = 'U/Y';
}
$sheet->setCellValue('H25', $um);

$sheet->setCellValue('G27', $potem30["toaddr"]["a16"]);
$sheet->setCellValue('H27', $potem30["toaddr"]["a17"]);

$sheet->setCellValue('A28', 'Total   Amount  ：'.$potem30["toaddr"]["a18"]);
$sheet->setCellValue('C28', $potem30["toaddr"]["a19"]);
$sheet->getStyle('C28')->getAlignment()->setWrapText(true);

$sheet->setCellValue('A29', 'Payment  Terms：'.$potem30["toaddr"]["a20"]);
$sheet->setCellValue('A30', 'Price   Terms    ：'.$potem30["toaddr"]["a21"]);


/**
 * 底部remark
 */
$sheet->setCellValue('A32', $potem30["remark"]["c2"][0]);
$sheet->setCellValue('B32', 'AMOUNT & QUANTITY WITHIN THE TOLERANCE OF '.$potem30["remark"]["c2"][1].'MORE OR LESS IS ONLY ALLOWED.');

$sheet->setCellValue('A33', $potem30["remark"]["c3"][0]);
$sheet->setCellValue('B33', 'YOUR PARTY MUST TAKE FULL RESPONSIBILITY FOR ANY DELAY OF SHIPMENT.');

$sheet->setCellValue('A34', $potem30["remark"]["c4"][0]);
$sheet->setCellValue('B34', 'AZO FREE');

$sheet->setCellValue('A35', $potem30["remark"]["c5"][0]);
$sheet->setCellValue('B35', 'ALL PERFORMANCES SHOULD MEET OUR REQUIREMENTS（AS PER ATTACHED）.');

$sheet->setCellValue('A36', $potem30["remark"]["c6"][0]);
$sheet->setCellValue('B36', 'YOU HAVE TO SUBMIT '.$potem30["remark"]["c6"][1].' SHIPMENT SAMPLE FOR OUR APPROVAL BEFORE '. $potem30["remark"]["c6"][2] .'OF SHIPMENT.');




$c72 =  $potem30["remark"]["c7"][2] ? ' EXCLUDING' : ' INCLUDING ' ;
$c73 =  $potem30["remark"]["c7"][3] ? ' TEST CHARGES ' : ' SURCHARGE ' ;


if($potem30["remark"]["c7"][1]){
    $c7value = $c72 . $c73;
    $sheet->setCellValue('A37', $potem30["remark"]["c7"][0]);
    $sheet->setCellValue('B37', $c7value);
}else{
    $c7value = '';
}



$sheet->setCellValue('B40', 'PLEASE CONFIRM AND COUNTER-SIGN BY RETURN. OTHERWISE, IF WE DO NOT RECEIVE ANY CONTRARY REPLIED WITHIN  '.$potem30["remark"]["c11"].',  THIS CONTRACT IS VALID.');
$sheet->getStyle('B40')->getAlignment()->setWrapText(true);

if($potem30["remark"]["c12"] == 1){
    $c12 = 'EXCLUDING';
}else{
    $c12 = 'INCLUDING';
}
$sheet->setCellValue('B42', $c12.' VAT INVOICE');
$sheet->setCellValue('B43', 'ORDER NO '.$potem30["remark"]["c13"]);

$row = 44;
if(is_array($potem30["remark"]["c14"])){
    if(count($potem30["remark"]["c14"]) > 1){
        foreach ($potem30["remark"]["c14"] as $item=>$value){

            if($item >1){
                $sheet->insertNewRowBefore($row, 1);
            }

            $sheet->setCellValue('A'. $row, $value );
            //$sheet->getStyle('A'. $row)->applyFromArray($noborderLeft);
            //$sheet->getStyle('A'. $row)->getAlignment()->setWrapText(true);

            $sheet->mergeCells("B{$row}:H{$row}");
            $sheet->setCellValue('B'. $row, $potem30["remark"]["c15"][$item]);
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
if(is_array($potem30["remark"]["c8"])){
    if(count($potem30["remark"]["c8"]) > 1){
        foreach ($potem30["remark"]["c8"] as $item=>$value){

            if($item >0){
                $sheet->insertNewRowBefore($row, 1);
            }

            $sheet->setCellValue('A'. $row, $value );
            //$sheet->getStyle('A'. $row)->applyFromArray($noborderLeft);


            $sheet->mergeCells("B{$row}:H{$row}");
            $sheet->setCellValue('B'. $row, $potem30["remark"]["c9"][$item]);
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
if(count($potem30["toaddr"]["a12"]) > 0){
    foreach ($potem30["toaddr"]["a12"] as $item=>$value){

        if($item >0){
            $sheet->insertNewRowBefore($row, 1);
        }

        $sheet->setCellValue('B'. $row, $value);
        $sheet->setCellValue('F'. $row, $potem30["toaddr"]["a13"][$item]);
        $sheet->setCellValue('G'. $row, $potem30["toaddr"]["a14"][$item]);
        $sheet->setCellValue('H'. $row, $potem30["toaddr"]["a15"][$item]);

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
if($potem30["orderform"]["elist"]['elistrow'] > 1){
    foreach ($potem30["orderform"]["elist"]['e1'] as $item=>$value){

        $sheet->insertNewRowBefore($row, 1);
        $sheet->mergeCells("B{$row}:F{$row}");
        $sheet->setCellValue('B'. $row, $value);
        $row++;
    }
}

/**
 * PO详情 附加
 */


/**
 *   以上为 主要内容
 */


$sheet->getPageSetup()
    ->setOrientation(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_PORTRAIT);  //竖放置
$sheet->getPageSetup()
    ->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);  //A4
$sheet->getPageSetup()->setFitToPage(true); //将工作表调整为一页

unset($_SESSION['potem30'] ); //注销SESSION

$filenameout = 'PO_'.$potem30['pono'];
outExcel($spreadsheet,$filenameout);


