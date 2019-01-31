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
fill_cell($styleArray, 'A1', 'A1', 'Sample order no.：');
fill_cell($styleArrayLefttop, 'A2', 'A2', 'Sketch：');
fill_cell($styleArray, 'A3', 'A3', '');

$remarkNums = array();

for ($col = 0; $col < count($cpsform['id']); $col++) {
    $Brow = chr(66 + $col * 2); //B
    $Crow = chr(67 + $col * 2); //C
    $BC   = "{$Brow}1:{$Crow}1";
    fill_cell($styleArray5, $BC, $Brow . '1', $cpsform['sampleorderno'][$col], $BC);

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
    $headerrow   = 4;
    foreach ($headertitle as $value) {
        fill_cell($styleArray, 'A' . $headerrow, 'A' . $headerrow, $value);
        $headerrow++;
    }

    $BC = "{$Brow}4:{$Crow}4";
    fill_cell($styleArray5, $BC, $Brow . '4', $cpsform['jobno'][$col], $BC);
    $BC = "{$Brow}5:{$Crow}5";
    fill_cell($styleArray5, $BC, $Brow . '5', $cpsform['styleno'][$col], $BC);
    $BC = "{$Brow}6:{$Crow}6";
    fill_cell($styleArray5, $BC, $Brow . '6', $cpsform['shipmentdate'][$col], $BC);
    fill_cell($styleArray5, $Brow . '7', $Brow . '7', $cpsform['alist'][$col]['a3'][0]);
    fill_cell($styleArray5, $Brow . '8', $Brow . '8', $cpsform['alist'][$col]['a5'][0]);
    $BC = "{$Crow}7:{$Crow}8";
    fill_cell($styleArray5, $BC, $Crow . '7', $cpsform['alist'][$col]['a4'][0], $BC);

    $img = $cpsform['alist'][$col]['a6'][0];
    fill_img($img, $Crow . '2', 180, 200);
    $spreadsheet->getActiveSheet()->getStyle($Crow . '2')->applyFromArray($styleArray); //设置边框

    if ($cpsform['shipmentlist'][$col][0] > 0) {
        $thisrow = 2;
        $smb     = '';
        for ($u = 0, $i = 1; $u < $cpsform['shipmentlist'][$col][0]; $u++, $i++) {
            if ($cpsform['shipmentlist'][$col]['sma' . $i] == 'on') {
                if ($u > 0) {
//                    $smb .= '
//'; //输出换行
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

    if ($cpsform['falist'][$col]['fa2alist'][0] > 0) {
        $fab2titlearr = $cpsform['titlearr']['fab2titlearr'];
        $thisrow      = 9;
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
                fill_cell($styleArray, 'A' . $thisrow, 'A' . $thisrow, $titlearr[$u]);
            }
            fill_cell($styleArray5, $Brow . $thisrow, $Brow . $thisrow, $cpsform["blist"][$col]["b" . $i][0]);
            fill_cell($styleArray5, $Crow . $thisrow, $Crow . $thisrow, $cpsform["clist"][$col]["c" . $i][0]);
            $thisrow++;
        }
    }
}

/**
 * shell
 */

for ($col = 0; $col < count($cpsform['id']); $col++) {
    $Brow = chr(66 + $col * 2); //B
    $Crow = chr(67 + $col * 2); //C
    if (isset($cpsform["falist"][$col]['falist']['fabrow']) && $cpsform["falist"][$col]['falist']['fabrow'] > 0) {
        $thisrow = 9;
        for ($u = 0, $i = 1; $u < $cpsform["falist"][$col]['falist']['fabrow']; $u++, $i++) {
            if ($col == 0) {
                $spreadsheet->getActiveSheet()->insertNewRowBefore(9, 1);
            }
        }
        $thisrow = 9;
        for ($u = 0, $i = 1; $u < $cpsform["falist"][$col]['falist']['fabrow']; $u++, $i++) {
            $fabtitlearr = $cpsform['titlearr']['fabtitlearr'];
            fill_cell($styleArray, 'A' . $thisrow, 'A' . $thisrow, $fabtitlearr[$u]);
            $BC = "{$Brow}{$thisrow}:{$Crow}{$thisrow}";
            fill_cell($styleArray5, $BC, $Brow . $thisrow, $cpsform["falist"][$col]['falist']["fa" . $i], $BC);
            $thisrow++;
        }
    }
}

 unset($_SESSION['cpsform']); //注销SESSION

foreach (range('B','M') as $item){
    for($i=1;$i<=100;$i++){
        $spreadsheet->getActiveSheet()->getStyle($item.$i)->getFont()->setSize(10);  //自动列宽度
    }

}

set_print_pcs('B6');

$spreadsheet->getActiveSheet()->getPageMargins()->setRight(0.1); //设置打印边距
$spreadsheet->getActiveSheet()->getPageMargins()->setLeft(0.1); //*/
set_writer($type);
