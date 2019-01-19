<?php

$type = 'HA';
// header
$xlsxName = 'product7all';
require_once 'common-header.php';

// content
$productall = $_SESSION['productall'];
$productp1  = $productall['productp1'];
$productp2  = $productall['productp2'];

/**
 * Page 1
 */
// 客户
fill_cell(null, 'B2', 'B2', $productp1['guest']);
// 开单日期
fill_cell(null, 'C3', 'C3', $productp1['alist']['a1']);
// 订单数量
fill_cell(null, 'G2', 'G2', $productp1['alist']['a11']);
// 制单号
fill_cell(null, 'L2', 'L2', $productp1['jobno']);
// 款号
fill_cell(null, 'L3', 'L3', $productp1['styleno']);
// 办单号
fill_cell(null, 'R2', 'R2', $productp1['sampleno']);
// 落货日期
fill_cell(null, 'R3', 'R3', $productp1['alist']['a2']);

//  工艺说明
fill_cell(null, 'A33:U43', 'A33', $productp1['fablist']['fab5']);
//  评语
fill_cell(null, 'A45:U50', 'A45', $productp1['fablist']['fab6']);
//  评语附加
fill_cell(null, 'A51:U56', 'A51', $productp1['fablist']['fab7']);
//  制单人
fill_cell(null, 'B57', 'B57', '制单人:  ' . $productp1['alist']['a4']);

//  特殊工艺
if ($productp1['alist']['a5value'] != null) {
    $teshu = join(",", $productp1['alist']['a5value']);
    fill_cell(null, 'R17:U17', 'R17', $teshu, 'R17:U17');
}
//  加工廠
fill_cell(null, 'R18:U18', 'R18', $productp1['ctlist']['ct14'], 'R18:U18');

fill_cell(null, 'P7', 'P7', $productp1['fablist']['fab2']);

//  裁法
$M13 = '单方向 ' . isselect($productp1['alist']['a6']) . ' 倒插 ' . isselect($productp1['alist']['a7']) . ' 女装 ' . isselect($productp1['alist']['a8']) . ' 不可倒插 ' . isselect($productp1['alist']['a10']);
fill_cell(null, 'P13:U14', 'P13', $M13);

$a9Yes = ($productp1['alist']['a9'] != null && $productp1['alist']['a9'] == "1") ? isselect('on') : isselect("off");
$a9No  = ($productp1['alist']['a9'] != null && $productp1['alist']['a9'] == "2") ? isselect('on') : isselect("off");
$P15   = '过粘朴机 ：' . ' 是 ' . $a9Yes . ' 否 ' . $a9No;
fill_cell($styleArray, 'P15:U15', 'P15', $P15, 'P15:U15');
fill_cell(null, 'P30:U30', 'P30', $productp1['fablist']['fab4']);

//  图片模块
$img = $productp1["alist"]["a3"];
fill_img($img, 'S4', 150, 150);

$row = 21;
fill_cell(null, 'D' . $row, 'D' . $row, $productp1['fablist']['bfablist']['bfablistremark']);
$row++;
for ($i = 0; $i < count($productp1['fablist']['bfablist']['bfab1']); $i++) {
    if ($i < 8) {
        if ($i == 0) {
            fill_cell(null, 'A' . $row, 'A' . $row, $productp1['fablist']['bfablist']['bfab1'][$i] . " " . $productp1['fablist']['bfablist']['bfab2'][$i]);
        } else {
            fill_cell(null, 'A' . $row, 'A' . $row, $productp1['fablist']['bfablist']['bfab1'][$i]);
            fill_cell(null, 'B' . $row, 'B' . $row, $productp1['fablist']['bfablist']['bfab2'][$i], 'B' . $row . ':N' . $row);
        }
    } else {
        $spreadsheet->getActiveSheet()->insertNewRowBefore($row, 1);
        fill_cell(null, 'A' . $row, 'A' . $row, $productp1['fablist']['bfablist']['bfab1'][$i]);
        fill_cell(null, 'B' . $row, 'B' . $row, $productp1['fablist']['bfablist']['bfab2'][$i], 'B' . $row . ':N' . $row);
    }
    $row++;
}

$row = 16;
for ($i = 0; $i < count($productp1['fablist']['afablist']['afab1']); $i++) {
    if ($i >= 5) {
        $spreadsheet->getActiveSheet()->insertNewRowBefore($row, 1);
    }
    fill_cell(null, 'A' . $row, 'A' . $row, $productp1['fablist']['afablist']['afab1'][$i]);
    fill_cell(null, 'B' . $row, 'B' . $row, $productp1['fablist']['afablist']['afab2'][$i], 'B' . $row . ':N' . $row);
    $row++;
}

//  细数分配表
for ($v = 0, $ct = 15; $v < 8; $v++, $ct++) {
    $col = chr(69 + $v); //D
    fill_cell($styleArray, $col . '5', $col . '5', $productp1['ctlist']['ct' . $ct]);
}

//  总计
$row  = 8;
$last = $productp1['allot']['formnum'];
for ($h = 3; $h <= 11; $h++) {
    if ($h == 11) {
        $col = chr(67 + $h);
    } else {
        $col = chr(66 + $h);
    }
    fill_cell($styleArray, $col . $row, $col . $row, $productp1['allot']['b' . $h][$last]);
}

//  总行数
$row = 7;
for ($i = 0; $i < $productp1['allot']['formnum']; $i++) {
    if ($i > 0) {
        $spreadsheet->getActiveSheet()->insertNewRowBefore($row, 1);
    }
    for ($h = 1; $h <= 11; $h++) {
        $merge = 'A' . $row . ":" . 'B' . $row;
        if ($h == 1) {
            $col = chr(64 + $h);
        } else if ($h == 11) {
            $col   = chr(67 + $h);
            $merge = 'N' . $row . ":" . 'O' . $row;
        } else {
            $col = chr(66 + $h);
        }
        fill_cell($styleArray, $col . $row, $col . $row, $productp1['allot']['b' . $h][$i], $merge);
    }
    $row++;
}
set_horizontal(false);

/**
 * Page 2
 */
$spreadsheet->setActiveSheetIndex(1);
fill_cell(null, 'C2', 'C2', $productp2['guest']);
fill_cell(null, 'C3', 'C3', $productp2['alist']['a1']);
fill_cell(null, 'J2', 'J2', $productp1['jobno'], 'J2:M2');
fill_cell(null, 'J3', 'J3', $productp2['styleno'], 'J3:M3');
fill_cell(null, 'R2', 'R2', $productp2['alist']['a11'], 'R2:S2');
fill_cell(null, 'R3', 'R3', $productp2['alist']['a2']);

// 主唛位置：
fill_cell($styleArray4, 'I7', 'I7', $productp2['blist']['b1'], 'I7:M7');
fill_cell($styleArray4, 'I12', 'I12', $productp2['blist']['b2'], 'I12:M12');
fill_cell($styleArray4, 'I15', 'I15', $productp2['blist']['b5'], 'I15:M15');
fill_img($productp2['blist']['b3'], 'B7', 250, 170);

// 洗水唛位置
fill_cell($styleArray4, 'C37', 'C37', $productp2['blist']['b6'], 'C37:L37');
fill_cell($styleArray4, 'C38', 'C38', $productp2['blist']['b7'], 'C38:L38');
fill_cell($styleArray4, 'C39', 'C39', $productp2['blist']['b8'], 'C39:L39');
fill_cell($styleArray4, 'C40', 'C40', $productp2['blist']['b9'], 'C40:L40');
fill_cell($styleArray4, 'D41', 'D41', $productp2['blist']['b10'], 'D41:L41');
fill_cell($styleArray4, 'C42', 'C42', $productp2['blist']['b11'], 'C42:L42');
fill_cell($styleArray4, 'C43', 'C43', $productp2['blist']['b12'], 'C43:L43');
fill_img($productp2['blist']['b4'], 'N34', 300, 420);

// 物料
$row = 23;
for ($i = 0; $i < count($productp1['fablist']['bfablist']['bfab1']); $i++) {
    if ($i < 6) {
        if ($i == 0) {
            fill_cell(null, 'A' . $row, 'A' . $row, $productp1['fablist']['bfablist']['bfab1'][$i] . " " . $productp1['fablist']['bfablist']['bfab2'][$i]);
        } else {
            fill_cell(null, 'A' . $row, 'A' . $row, $productp1['fablist']['bfablist']['bfab1'][$i], 'A' . $row . ':B' . $row);
            fill_cell($styleArray4, 'C' . $row, 'C' . $row, $productp1['fablist']['bfablist']['bfab2'][$i], 'C' . $row . ':T' . $row);
        }
    } else {
        $spreadsheet->getActiveSheet()->insertNewRowBefore($row, 1);
        fill_cell(null, 'A' . $row, 'A' . $row, $productp1['fablist']['bfablist']['bfab1'][$i], 'A' . $row . ':B' . $row);
        fill_cell($styleArray4, 'C' . $row, 'C' . $row, $productp1['fablist']['bfablist']['bfab2'][$i], 'C' . $row . ':T' . $row);
    }
    $row++;
}

set_horizontal(false);

/**
 * Page 3
 */
set_ha_p3();

/**
 * Page 4
 */
set_ha_p4();

// footer
// unset($_SESSION['productall']); //注销SESSION
set_writer($type);
