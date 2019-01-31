<?php

require_once 'aidenfunc.php';

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


//填数据
$spreadsheet->getActiveSheet()->setCellValue('A9', $potem30["tosb"]);
$spreadsheet->getActiveSheet()->setCellValue('B18', $potem30["podate"]);

//if(1 == $potem30["toaddr"]["a1"]){
//  $titlecon = 'HIGH ABLE INVESTMENT LIMITED';
//}elseif (2 == $potem30["toaddr"]["a1"]){
//    $titlecon = 'IRONDALE FASHION INTERNATIONAL LIMITED';
//}
//$spreadsheet->getActiveSheet()->setCellValue('A1', $titlecon);
setCell($sheet,"A1",$potem30["remark"]['poheader']['poheada1'],$noborderCenter);
setCell($sheet,"A2",$potem30["remark"]['poheader']['poheada2'],$noborderCenter);
setCell($sheet,"A3",$potem30["remark"]['poheader']['poheada3'],$noborderCenter);
setCell($sheet,"A4",'Tel:'.$potem30["remark"]['poheader']['poheada5'].'Fax:'.$potem30["remark"]['poheader']['poheada5'],$noborderCenter);
//setCell($sheet,"A5",'Attn :'.$potem1["toaddr"]["a9"],$noborderCenter);



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

$spreadsheet->getActiveSheet()->setCellValue('G27', $potem30["toaddr"]["a16"]);
$spreadsheet->getActiveSheet()->setCellValue('H27', $potem30["toaddr"]["a17"]);

$spreadsheet->getActiveSheet()->setCellValue('A28', 'Total   Amount  ：'.$potem30["toaddr"]["a18"]);
$spreadsheet->getActiveSheet()->setCellValue('C28', $potem30["toaddr"]["a19"]);
$spreadsheet->getActiveSheet()->getStyle('C28')->getAlignment()->setWrapText(true);

$spreadsheet->getActiveSheet()->setCellValue('A29', 'Payment  Terms：'.$potem30["toaddr"]["a20"]);
$spreadsheet->getActiveSheet()->setCellValue('A30', 'Price   Terms    ：'.$potem30["toaddr"]["a21"]);


/**
 * 底部remark
 */
$spreadsheet->getActiveSheet()->setCellValue('A32', $potem30["remark"]["c2"][0]);
$spreadsheet->getActiveSheet()->setCellValue('B32', 'AMOUNT & QUANTITY WITHIN THE TOLERANCE OF '.$potem30["remark"]["c2"][1].'MORE OR LESS IS ONLY ALLOWED.');

$spreadsheet->getActiveSheet()->setCellValue('A33', $potem30["remark"]["c3"][0]);
$spreadsheet->getActiveSheet()->setCellValue('B33', 'YOUR PARTY MUST TAKE FULL RESPONSIBILITY FOR ANY DELAY OF SHIPMENT.');

$spreadsheet->getActiveSheet()->setCellValue('A34', $potem30["remark"]["c4"][0]);
$spreadsheet->getActiveSheet()->setCellValue('B34', 'AZO FREE');

$spreadsheet->getActiveSheet()->setCellValue('A35', $potem30["remark"]["c5"][0]);
$spreadsheet->getActiveSheet()->setCellValue('B35', 'ALL PERFORMANCES SHOULD MEET OUR REQUIREMENTS（AS PER ATTACHED）.');

$spreadsheet->getActiveSheet()->setCellValue('A36', $potem30["remark"]["c6"][0]);
$spreadsheet->getActiveSheet()->setCellValue('B36', 'YOU HAVE TO SUBMIT '.$potem30["remark"]["c6"][1].' SHIPMENT SAMPLE FOR OUR APPROVAL BEFORE '. $potem30["remark"]["c6"][2] .'OF SHIPMENT.');




$c72 =  $potem30["remark"]["c7"][2] ? ' EXCLUDING' : ' INCLUDING ' ;
$c73 =  $potem30["remark"]["c7"][3] ? ' TEST CHARGES ' : ' SURCHARGE ' ;


if($potem30["remark"]["c7"][1]){
    $c7value = $c72 . $c73;
    $spreadsheet->getActiveSheet()->setCellValue('A37', $potem30["remark"]["c7"][0]);
    $spreadsheet->getActiveSheet()->setCellValue('B37', $c7value);
}else{
    $c7value = '';
}



$spreadsheet->getActiveSheet()->setCellValue('B40', 'PLEASE CONFIRM AND COUNTER-SIGN BY RETURN. OTHERWISE, IF WE DO NOT RECEIVE ANY CONTRARY REPLIED WITHIN  '.$potem30["remark"]["c11"].',  THIS CONTRACT IS VALID.');
$spreadsheet->getActiveSheet()->getStyle('B40')->getAlignment()->setWrapText(true);

if($potem30["remark"]["c12"] == 1){
    $c12 = 'EXCLUDING';
}else{
    $c12 = 'INCLUDING';
}
$spreadsheet->getActiveSheet()->setCellValue('B42', $c12.' VAT INVOICE');
$spreadsheet->getActiveSheet()->setCellValue('B43', 'ORDER NO '.$potem30["remark"]["c13"]);

$row = 44;
if(is_array($potem30["remark"]["c14"])){
    if(count($potem30["remark"]["c14"]) > 1){
        foreach ($potem30["remark"]["c14"] as $item=>$value){

            if($item >1){
                $spreadsheet->getActiveSheet()->insertNewRowBefore($row, 1);
            }

            $sheet->setCellValue('A'. $row, $value );
            //$spreadsheet->getActiveSheet()->getStyle('A'. $row)->applyFromArray($noborderLeft);
            //$spreadsheet->getActiveSheet()->getStyle('A'. $row)->getAlignment()->setWrapText(true);

            $spreadsheet->getActiveSheet()->mergeCells("B{$row}:H{$row}");
            $sheet->setCellValue('B'. $row, $potem30["remark"]["c15"][$item]);
            $spreadsheet->getActiveSheet()->getStyle('B'. $row)->applyFromArray($noborderLeft);
            $spreadsheet->getActiveSheet()->getStyle('B'. $row)->getAlignment()->setWrapText(true);

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
                $spreadsheet->getActiveSheet()->insertNewRowBefore($row, 1);
            }

            $sheet->setCellValue('A'. $row, $value );
            //$spreadsheet->getActiveSheet()->getStyle('A'. $row)->applyFromArray($noborderLeft);


            $spreadsheet->getActiveSheet()->mergeCells("B{$row}:H{$row}");
            $sheet->setCellValue('B'. $row, $potem30["remark"]["c9"][$item]);
            $spreadsheet->getActiveSheet()->getStyle('B'. $row)->applyFromArray($noborderLeft);
            $spreadsheet->getActiveSheet()->getStyle('B'. $row)->getAlignment()->setWrapText(true);

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
            $spreadsheet->getActiveSheet()->insertNewRowBefore($row, 1);
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

        $spreadsheet->getActiveSheet()->insertNewRowBefore($row, 1);
        $spreadsheet->getActiveSheet()->mergeCells("B{$row}:F{$row}");
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



$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页

//unset($_SESSION['potem30'] ); //注销SESSION

$filenameout = 'PO_'.$potem30['shortName'];
outExcel($spreadsheet,$filenameout);


