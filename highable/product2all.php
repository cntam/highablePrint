<?php
$type = 'HA';
// header
$xlsxName = 'product2all';
require_once 'common-header.php';

// content
$productall = $_SESSION['productall'];
$productp1  = $productall['productp1'];
$productp2  = $productall['productp2'];

/**
 * Page 1
 */
$col = 'A';
for ($i=0 ;$i<=17;$i++){
    $spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(6);  //列宽度
    $col++;
}
//
// 客户
fill_cell(null, 'C2', 'C2', $productp1['guest']);
// 开单日期
fill_cell(null, 'C3', 'C3', $productp1['alist']['a1']);
// 制单号
fill_cell(null, 'I2', 'I2', $productp1['jobno']);
// 款号
fill_cell(null, 'I3', 'I3', $productp1['styleno']);
// 办单号
fill_cell(null, 'N2', 'N2', $productp1['sampleno']);
// 落货日期
fill_cell(null, 'O3', 'O3', $productp1['alist']['a2']);

//  船头办数量
//  底部附加行 remark
fill_cell($styleArray, 'E15:L15', 'E15', $productp1['ctlist']['ct23']);

for ($i = 0, $ct = 1; $i < 14; $i++, $ct++) {
    if ($ct == 14) {
        fill_cell($styleArray, 'A16', 'A16', '出船头办日期：' . $productp1['ctlist']['ct' . $ct]);
    } elseif ($ct == 12 or $ct == 13) {
        if ($ct == 12) {
            //  净重：
            fill_cell($styleArray, 'C15', 'A15', '净重：' . $productp1['ctlist']['ct' . $ct]);
        } else {
            //  毛重：
            fill_cell($styleArray, 'F15', 'C15', '毛重：' . $productp1['ctlist']['ct' . $ct]);
        }
    } else {
        $row = 14;
        if ($ct == 1) {
            $col = chr(64 + $ct); //A
        } else {
            $col = chr(64 + $ct); //B
        }
        fill_cell($styleArray, $col . $row, $col . $row, $productp1['ctlist']['ct' . $ct]);
    }
}

//  生产办
// fill_cell( $styleArray, 'A17', 'A17', '生产办：' . $productp1['ctlist']['ct24']);

//  工艺说明
fill_cell($styleArraytop, 'A22:R28', 'A22', $productp1['fablist']['fab5']);
//  评语
fill_cell($styleArraytop, 'A30:R53', 'A30', $productp1['fablist']['fab6']);
//  评语附加
fill_cell($styleArraytop, 'A54:R59', 'A54', $productp1['fablist']['fab7']);
//  制单人
fill_cell(null, $col . $row, 'B60', '制单人:  ' . $productp1['alist']['a4']);

//  特殊工艺
if ($productp1['alist']['a5value'] != null) {
    $teshu = join(",", $productp1['alist']['a5value']);
    fill_cell(null, 'O5:P5', 'O5', $teshu);
}

fill_cell(null, 'Q4:R14', 'Q4', $productp1['fablist']['fab2']);

//  裁法
$M16 = '单方向：' . isselect($productp1['alist']['a6']) . '倒插：' . isselect($productp1['alist']['a7']) . '女装：' . isselect($productp1['alist']['a8']);
fill_cell(null, 'M16:R16', 'M16', $M16);

fill_cell(null, 'M17:R19', 'M17', $productp1['fablist']['fab3']);

$a9Yes = ($productp1['alist']['a9'] != null && $productp1['alist']['a9'] == "1") ? isselect('on') : isselect("off");
$a9No  = ($productp1['alist']['a9'] != null && $productp1['alist']['a9'] == "2") ? isselect('on') : isselect("off");
$M19   = '过粘朴机 ：' . '是：' . $a9Yes . ' 否：' . $a9No;
fill_cell(null, 'M19:R19', 'M19', $M19);
fill_cell(null, 'M20:R20', 'M20', '针距：' . $productp1['fablist']['fab4']);

//  图片模块
$img = $productp1["alist"]["a3"];
fill_img($img, 'M7', 180, 300);

//  细数分配表
for ($v = 0, $ct = 15; $v < 8; $v++, $ct++) {
    $col = chr(68 + $v); //D
    fill_cell($styleArray, $col . '5', $col . '5', $productp1['ctlist']['ct' . $ct]);
}

//  总计
$row  = 12;
$last = $productp1['allot']['formnum'];
for ($h = 3; $h <= 11; $h++) {
    $col = chr(65 + $h); //B
    fill_cell($styleArray, $col . $row, $col . $row, $productp1['allot']['b' . $h][$last]);
}

//  总行数
$row = 7;
for ($i = 0; $i < $productp1['allot']['formnum']; $i++) {
    if ($i > 4) {
        $spreadsheet->getActiveSheet()->insertNewRowBefore($row, 1);
        for ($h = 1; $h <= 11; $h++) {
            if ($h == 1) {
                $col = chr(64 + $h); //A
            } else {
                $col = chr(65 + $h); //B
            }
            fill_cell($styleArray, $col . $row, $col . $row, $productp1['allot']['b' . $h][$i], 'A' . $row . ":" . 'B' . $row);
        }
        $row++;
    } else {
        for ($h = 1; $h <= 11; $h++) {
            if ($h == 1) {
                $col = chr(64 + $h); //A
            } else {
                $col = chr(65 + $h); //B
            }
            fill_cell($styleArray, $col . $row, $col . $row, $productp1['allot']['b' . $h][$i]);
        }
        $row++;
    }
}
//set_horizontal(false);

set_horizontal(false, true);
//// 所有行打印在一页
//set_fitToWidth(0);

/**
 * Page 2
 */
$spreadsheet->setActiveSheetIndex(1);
$col = 'A';
for ($i=0 ;$i<=17;$i++){
    $spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(6);  //列宽度
    $col++;
}
fill_cell(null, 'C2', 'C2', $productp2['guest']);
fill_cell(null, 'C3', 'C3', $productp2['alist']['a1']);
fill_cell(null, 'I2', 'I2', $productp2['jobno']);
fill_cell(null, 'I3', 'I3', $productp2['styleno']);
fill_cell(null, 'M2', 'M2', $productp2['sampleno']);
fill_cell(null, 'N3', 'N3', $productp2['alist']['a2']);

// 主唛/烟治唛/产地唛车法：
fill_cell(null, 'F7', 'F7', $productp2['blist']['b1']);
fill_cell(null, 'F17', 'F17', $productp2['blist']['b2']);
fill_cell(null, 'C18', 'C18', $productp2['blist']['b9']);
fill_img($productp2['blist']['b3'], 'F8', 250, 170);

// 洗水唛位置
fill_cell(null, 'E40', 'E40', $productp2['blist']['b5']);
fill_cell(null, 'P38', 'P38', $productp2['blist']['b6']);
fill_img($productp2['blist']['b4'], 'F26', 400, 300);

// 挂牌位置
fill_cell(null, 'A43', 'A43', $productp2['blist']['b8']);
fill_img($productp2['blist']['b7'], 'D44', 250, 500);
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
