<?php
$type = "CPS";
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
if ($cpsform['titlearr']['temno'] == 3) {
    fill_cell($styleArray, 'A1', 'A1', 'LAUK PCS');
}

fill_cell($styleArrayLefttop, 'A2', 'A2', 'Sketch：');

fill_cell($styleArray, 'A3', 'A3', '');

$remarkNums = array();

for ($col = 0; $col < count($cpsform['id']); $col++) {
    $Brow = chr(66 + $col * 2); //B
    $Crow = chr(67 + $col * 2); //C

    /**
     *  外厂 及大货是否完成
     */
    if ($cpsform['alist'][$col]['a7'][0] == 'on') {
        $completeIcon = '  completed';
    } else {
        $completeIcon = '';
    }

    fill_cell($styleArray5, $Brow . '3', $Brow . '3', $cpsform['alist'][$col]['a1'][0] . $completeIcon, $Brow . '3:' . $Crow . '3');

    $BC = "{$Brow}3:{$Crow}3";
    fill_cell($styleArray5, $BC, $Crow . '3', $cpsform['alist'][$col]['a2'][0]);

    $headertitle = $cpsform['titlearr']['headertitle'];
    $headerrow   = 4;
    foreach ($headertitle as $value) {
        fill_cell($styleArray, 'A' . $headerrow, 'A' . $headerrow, $value);
        $headerrow++;
    }

    if ($cpsform['titlearr']['temno'] == 3) {
        $BC = "{$Brow}4:{$Crow}4";
        fill_cell($styleArray5, $BC, $Brow . '4', $cpsform['sampleorderno'][$col], $BC);

        $BC = "{$Brow}5:{$Crow}5";
        fill_cell($styleArray5, $BC, $Brow . '5', $cpsform['jobno'][$col], $BC);

        $BC = "{$Brow}6:{$Crow}6";
        fill_cell($styleArray5, $BC, $Brow . '6', $cpsform['styleno'][$col], $BC);
        $BC = "{$Brow}7:{$Crow}7";
        fill_cell($styleArray5, $BC, $Brow . '7', $cpsform['shipmentdate'][$col], $BC);
        $BC = "{$Brow}8:{$Crow}8";
        fill_cell($styleArray5, $BC, $Brow . '8', $cpsform['alist'][$col]['a5'][0], $BC);

    } elseif ($cpsform['titlearr']['temno'] == 4) {
        $row = 4;
        $BC  = "{$Brow}{$row}:{$Crow}{$row}";
        fill_cell($styleArray5, $BC, $Brow . $row, $cpsform['jobno'][$col], $BC);
        $row++;

        $BC = "{$Brow}{$row}:{$Crow}{$row}";
        fill_cell($styleArray5, $BC, $Brow . $row, $cpsform['styleno'][$col], $BC);
        $row++;

        $BC = "{$Brow}{$row}:{$Crow}{$row}";
        fill_cell($styleArray5, $BC, $Brow . $row, $cpsform['shipmentdate'][$col], $BC);
        $row++;

        $BC = "{$Brow}{$row}:{$Crow}{$row}";
        fill_cell($styleArray5, $BC, $Brow . $row, $cpsform['alist'][$col]['a5'][0], $BC);
    }

    /**
     * 图片模块
     */
    $img = $cpsform['alist'][$col]['a6'][0];
    fill_img($img, $Crow . '2', 180, 200);
    $spreadsheet->getActiveSheet()->getStyle($Crow . '2')->applyFromArray($styleArray); //设置边框

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

    /**
     * fa2alist
     */
    if ($cpsform['falist'][$col]['fa2alist'][0] > 0) {
        $fab2titlearr = $cpsform['titlearr']['fab2titlearr'];
        if ($cpsform['titlearr']['temno'] == 3) {
            $thisrow = 9;
        } elseif ($cpsform['titlearr']['temno'] == 4) {
            $thisrow = 8;
        }
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
    $BC   = "{$Brow}{$thisrow}:{$Crow}{$thisrow}";
    if (count($cpsform['elist'][$col]['e1']) >= 0) {
        if ($cpsform['titlearr']['temno'] == 3) {
            $thisrow = 13;
        } elseif ($cpsform['titlearr']['temno'] == 4) {
            $thisrow = 11;
        }
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
        if ($cpsform['titlearr']['temno'] == 3) {
            $thisrow = 13;
        } elseif ($cpsform['titlearr']['temno'] == 4) {
            $thisrow = 11;
        }
        for ($u = 0, $i = 1; $u < count($titlearr); $u++, $i++) {
            if ($col == 0) {
                $spreadsheet->getActiveSheet()->insertNewRowBefore($thisrow, 1);
            }
            fill_cell($styleArray, 'A' . $thisrow, 'A' . $thisrow, $titlearr[$u]);
            $thisrow++;
        }

        if ($cpsform['titlearr']['temno'] == 3) {
            $thisrow = 13;
        } elseif ($cpsform['titlearr']['temno'] == 4) {
            $thisrow = 11;
        }
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
                if ($cpsform['titlearr']['temno'] == 3) {
                    $spreadsheet->getActiveSheet()->insertNewRowBefore(9, 1);
                } elseif ($cpsform['titlearr']['temno'] == 4) {
                    $spreadsheet->getActiveSheet()->insertNewRowBefore(8, 1);
                }
            }
        }
        // 主布料
        if ($cpsform['titlearr']['temno'] == 3) {
            $thisrow = 9;
        } elseif ($cpsform['titlearr']['temno'] == 4) {
            $thisrow = 8;
        }
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
 *  头部2栏 信息
 */
if ($cpsform['titlearr']['temno'] == 3) {
    $spreadsheet->getActiveSheet()->insertNewRowBefore(2, 1);
    setRowHeight('2', 20);
    for ($col = 0; $col < count($cpsform['id']); $col++) {
        $Brow = chr(66 + $col * 2); //B
        $Crow = chr(67 + $col * 2); //C
        $row  = 1;
        fill_cell($styleArray5, 'A' . $row, 'A' . $row, 'LAUK CPS');
        $a = 8;
        for ($y = 1; $y <= 2; $y++) {
            fill_cell($styleArray5, $Brow . $row, $Brow . $row, $cpsform['alist'][$col]['a' . $a][0]);
            $a++;
            fill_cell($styleArray5, $Crow . $row, $Crow . $row, $cpsform['alist'][$col]['a' . $a][0]);
            $a++;
            $row++;
        }
    }
} elseif ($cpsform['titlearr']['temno'] == 4) {
    $spreadsheet->getActiveSheet()->insertNewRowBefore(2, 2);
    setRowHeight('2', 20);
    setRowHeight('3', 20);
    for ($col = 0; $col < count($cpsform['id']); $col++) {
        $Brow = chr(66 + $col * 2); //B
        $Crow = chr(67 + $col * 2); //C
        $row  = 1;
        fill_cell($styleArray5, 'A' . $row, 'A' . $row, 'LAJ');
        $a = 8;
        fill_cell($styleArray5, $Brow . $row, $Brow . $row, $cpsform['alist'][$col]['a' . $a][0]);
        $a++;
        fill_cell($styleArray5, $Crow . $row, $Crow . $row, $cpsform['alist'][$col]['a' . $a][0]);
        $a++;
        $row++;
        fill_cell($styleArray, 'A' . $row, 'A' . $row, '');
        $BC = "{$Brow}{$row}:{$Crow}{$row}";
        fill_cell($styleArray, $BC, $Brow . $row, $cpsform['alist'][$col]['a' . $a][0], $BC);
        $row++;
        fill_cell($styleArray, 'A' . $row, 'A' . $row, '');
        $BC = "{$Brow}{$row}:{$Crow}{$row}";
        fill_cell($styleArray6, $BC, $Brow . $row, $cpsform['sampleorderno'][$col], $BC);
    }
}

foreach (range('B','M') as $item){
    for($i=1;$i<=100;$i++){
        $spreadsheet->getActiveSheet()->getStyle($item.$i)->getFont()->setSize(10);  //自动列宽度
    }

}

set_print_pcs('B8');
$spreadsheet->getActiveSheet()->getPageMargins()->setRight(0.1); //设置打印边距
$spreadsheet->getActiveSheet()->getPageMargins()->setLeft(0.1); //*/
set_writer($type);
