<?php

$type = 'CPS';
require_once 'common-header.php';
$cpsform = $_SESSION['cpsform'];



for ($col = 0; $col < count($cpsform['id']); $col++) {
    $Brow = chr(66 + $col * 2); //B
    $Crow = chr(67 + $col * 2); //C
    setColumnWidth($Brow, 20);
    setColumnWidth($Crow, 20);
}

setRowHeight('1', 36);
setRowHeight('2', 160);

/**
 * 头部
 */
if ($cpsform['temno'][0] == "7") {
    fill_cell($styleArray, 'A1', 'A1', 'MCQUEEN');
} else {
    fill_cell($styleArray, 'A1', 'A1', 'Proenza Schouler');
}

fill_cell($styleArrayLefttop, 'A2', 'A2', 'Sketch：');

fill_cell($styleArray, 'A3', 'A3', '');

$remarkNums = array();

for ($col = 0; $col < count($cpsform['id']); $col++) {
    $Brow = chr(66 + $col * 2); //B
    $Crow = chr(67 + $col * 2); //C

    $BC = "{$Brow}1:{$Crow}1";
    fill_cell($styleArray5, $BC, $Brow . '1', $cpsform['sampleorderno'][$col], $BC);

    $img = $cpsform['alist'][$col]['a6'][0];
    fill_img($img, $Crow . '2', 180, 200);
    $spreadsheet->getActiveSheet()->getStyle($Crow . '2')->applyFromArray($styleArray); //设置边框

    /**
     *  外厂 及大货是否完成
     */

    if ($cpsform['alist'][$col]['a7'][0] == 'on') {
        $completeIcon = '  completed';
    } else {
        $completeIcon = '';
    }

    fill_cell($styleArray5, $Brow . '3', $Brow . '3', $cpsform['alist'][$col]['a1'][0] . $completeIcon);

    $BC = "{$Brow}3:{$Crow}3";
    fill_cell($styleArray5, $BC, $Crow . '3', $cpsform['alist'][$col]['a2'][0]);

    $headertitle = $cpsform['titlearr']['headertitle'];

    $BC = "{$Brow}4:{$Crow}4";
    fill_cell($styleArray, 'A4', 'A4', $headertitle[0]);
    fill_cell($styleArray5, $BC, $Brow . '4', $cpsform['jobno'][$col], $BC);

    $BC = "{$Brow}5:{$Crow}5";
    fill_cell($styleArray, 'A5', 'A5', $headertitle[1]);
    fill_cell($styleArray5, $BC, $Brow . '5', $cpsform['styleno'][$col], $BC);

    $BC = "{$Brow}6:{$Crow}6";
    fill_cell($styleArray, 'A6', 'A6', $headertitle[2]);
    fill_cell($styleArray5, $BC, $Brow . '6', $cpsform['sampleorderno'][$col], $BC);

    $BC = "{$Brow}7:{$Crow}7";
    fill_cell($styleArray, 'A7', 'A7', $headertitle[4]);
    if ($cpsform['temno'][0] == "7") {
        fill_cell($styleArray5, $BC, $Brow . '7', $cpsform['alist'][$col]['a5'][0], $BC);
    } else if ($cpsform['temno'][0] == "8") {
        fill_cell($styleArray5, $Brow . '7', $Brow . '7', $cpsform['alist'][$col]['a5'][0]);
        fill_cell($styleArray5, $Crow . '7', $Crow . '7', $cpsform['alist'][$col]['a4'][0]);
    }

    /**
     *  pcs
     */
    if ($cpsform['shipmentlist'][$col][0] > 0) {
        $thisrow = 2;
        $smb     = '';
        for ($u = 0, $i = 1; $u < $cpsform['shipmentlist'][$col][0]; $u++, $i++) {
            if ($cpsform['shipmentlist'][$col]['sma' . $i] == 'on') {
                if ($u > 0) {
//                    $smb .= '
//    '; //输出换行
                    $smb .= PHP_EOL; //输出换行
                }
                foreach ($cpsform['shipmentlist'][$col]['smb' . $i] as $item => $value) {
                    if ($item == 4) {
                        $smb .= ' ' . gmdate("d/M", strtotime($value));
                    } elseif($item == 0) {
                        $smb .= $value;
                    }else{
                        $smb .= ' ' . $value;
                    }
                }
                if ($cpsform['shipmentlist'][$col]['smc' . $i] == 'on') {
                    $smb .= '  已出货';
                }
            }
        }
        fill_cell($styleArray, $Brow . $thisrow, $Brow . $thisrow, $smb);
    }

    /**
     * fa2alist
     */
    if ($cpsform['falist'][$col]['fa2alist'][0] > 0) {
        $fab2titlearr = $cpsform['titlearr']['fab2titlearr'];
        $thisrow      = 8;
        for ($u = 0, $i = 1; $u < count($cpsform['falist'][$col]['fa2alist']['fa2a1']); $u++, $i++) {
            fill_cell($styleArray, 'A' . $thisrow, 'A' . $thisrow, stripcslashes($fab2titlearr[$u]));
            $BC = "{$Brow}{$thisrow}:{$Crow}{$thisrow}";
            fill_cell($styleArray5, $BC, $Brow . $thisrow, stripcslashes($cpsform['falist'][$col]['fa2alist']['fa2a1'][$u]), $BC);
            $thisrow++;
        }
    }

    array_push($remarkNums, count($cpsform['elist'][$col]['e1']));

}

/**
 * 底部remark
 */
for ($col = 0; $col < count($cpsform['id']); $col++) {
    $Brow = chr(66 + $col * 2); //B
    $Crow = chr(67 + $col * 2); //C
    if (count($cpsform['elist'][$col]['e1']) >= 0) {
        $thisrow = 11;
        for ($u = 0, $i = 1; $u < max($remarkNums); $u++, $i++) {
            if (count($cpsform['elist'][$col]['e1']) <= max($remarkNums)) {
                if (count($cpsform['elist'][$col]['e1']) == max($remarkNums)) {
                    fill_cell($styleArray, 'A' . $thisrow, 'A' . $thisrow, $cpsform['elist'][$col]['e1'][$u]);
                }
                fill_cell($styleArray5, $Brow . $thisrow, $Brow . $thisrow, $cpsform['elist'][$col]['e2'][$u]);
                fill_cell($styleArray5, $Crow . $thisrow, $Crow . $thisrow, $cpsform['elist'][$col]['e3'][$u]);
            } else {
                fill_cell($styleArray5, $Brow . $thisrow, $Brow . $thisrow, '');
                fill_cell($styleArray5, $Crow . $thisrow, $Crow . $thisrow, '');
            }
            $thisrow++;
        }
    }
}

/**
 *  Total Trim Cost row+
 */
for ($col = 0; $col < count($cpsform['id']); $col++) {
    $Brow = chr(66 + $col * 2); //B
    $Crow = chr(67 + $col * 2); //C
    if ($cpsform["blist"][$col][0] > 0) {
        $titlearr = $cpsform['titlearr']['titlearr'];
        $thisrow  = 11;
        for ($u = 0, $i = 1; $u < count($titlearr); $u++, $i++) {
            if ($col == 0) {
                $spreadsheet->getActiveSheet()->insertNewRowBefore($thisrow, 1);
            }
            fill_cell($styleArray, 'A' . $thisrow, 'A' . $thisrow, $titlearr[$u]);
            $thisrow++;
        }
        $thisrow = 11;
        for ($u = 0, $i = 1; $u < count($titlearr); $u++, $i++) {
            fill_cell($styleArray5, $Brow . $thisrow, $Brow . $thisrow, $cpsform["blist"][$col]["b" . $i][0]);
            fill_cell($styleArray5, $Crow . $thisrow, $Crow . $thisrow, $cpsform["clist"][$col]["c" . $i][0]);
            $thisrow++;
        }
    }
}

/**
 * shell 主布料
 */
for ($col = 0; $col < count($cpsform['id']); $col++) {
    $Brow = chr(66 + $col * 2); //B
    $Crow = chr(67 + $col * 2); //C
    if (isset($cpsform["falist"][$col]['falist']['fabrow']) && $cpsform["falist"][$col]['falist']['fabrow'] > 0) {
        for ($u = 0, $i = 1; $u < $cpsform["falist"][$col]['falist']['fabrow']; $u++, $i++) {
            if ($col == 0) {
                $spreadsheet->getActiveSheet()->insertNewRowBefore(8, 1);
            }
        }
        $thisrow = 8;
        for ($u = 0, $i = 1; $u < $cpsform["falist"][$col]['falist']['fabrow']; $u++, $i++) {
            $fabtitlearr = $cpsform['titlearr']['fabtitlearr'];
            fill_cell($styleArray, 'A' . $thisrow, 'A' . $thisrow, $fabtitlearr[$u]);
            $BC = "{$Brow}{$thisrow}:{$Crow}{$thisrow}";
            fill_cell($styleArray5, $BC, $Brow . $thisrow, $cpsform["falist"][$col]['falist']["fa" . $i], $BC);
            $thisrow++;
        }
    }
}

/**
 * shipment date
 */
for ($col = 0; $col < count($cpsform['id']); $col++) {
    $Brow = chr(66 + $col * 2); //B
    $Crow = chr(67 + $col * 2); //C
    $row  = 7;
    if ($cpsform['temno'][0] == "7") {
        if ($col == 0) {
            $spreadsheet->getActiveSheet()->insertNewRowBefore($row, 2);
            fill_cell(null, 'A7:A8', 'A' . $row, $cpsform['titlearr']['headertitle'][3], 'A7:A8');
        }
        fill_cell($styleArray5, $Brow . $row, $Brow . $row, $cpsform['shipmentdate'][$col]);
        fill_cell($styleArray5, $Crow . $row, $Crow . $row, $cpsform['alist'][$col]['a3'][0]);
        fill_cell($styleArray5, $Brow . ($row + 1), $Brow . ($row + 1), $cpsform['alist'][$col]['a4'][0], $Brow . ($row + 1) . ":" . $Crow . ($row + 1));
    } else if ($cpsform['temno'][0] == "8") {
        if ($col == 0) {
            $spreadsheet->getActiveSheet()->insertNewRowBefore($row, 1);
            fill_cell(null, 'A' . $row, 'A' . $row, $cpsform['titlearr']['headertitle'][3]);
        }
        fill_cell($styleArray5, $Brow . $row . ":" . $Crow . $row, $Brow . $row, $cpsform['shipmentdate'][$col], $Brow . $row . ":" . $Crow . $row);
    }
}

//$spreadsheet->getActiveSheet()->getStyle("A".$listrow)->getFont()->setSize(8);
foreach (range('B','M') as $item){
    for($i=1;$i<=100;$i++){
        $spreadsheet->getActiveSheet()->getStyle($item.$i)->getFont()->setSize(10);  //自动列宽度
    }

}

set_print_pcs('B6');
$spreadsheet->getActiveSheet()->getPageMargins()->setRight(0.1); //设置打印边距
$spreadsheet->getActiveSheet()->getPageMargins()->setLeft(0.1); //*/
set_writer($type);
