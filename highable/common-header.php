<?php
session_start();
error_reporting(0);
require_once 'autoloadconfig.php'; //判断是否在线
require_once 'img.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;

switch ($type) {
    case 'HA':
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/' . $xlsxName . '.xlsx');
        break;
    case 'CPS':
        $spreadsheet = new Spreadsheet();
        $spreadsheet->getActiveSheet()->setTitle("CPS");
        $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(20); //列宽度
        break;
    default:
        $spreadsheet = new Spreadsheet();
        break;
}

$spreadsheet->setActiveSheetIndex(0);
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft Yahei');
$spreadsheet->getDefaultStyle()->getFont()->setSize(12);

$styleArray = get_style(Alignment::HORIZONTAL_LEFT, Alignment::VERTICAL_CENTER, 1);

$borders = [
    'top' => [
        'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
    ],
];
$styleArraytop = get_style(Alignment::HORIZONTAL_LEFT, Alignment::VERTICAL_TOP, $borders);

$styleArrayLefttop = get_style(Alignment::HORIZONTAL_LEFT, Alignment::VERTICAL_TOP, 1);

$styleArray1 = get_style(Alignment::HORIZONTAL_LEFT, Alignment::VERTICAL_TOP, 1);

$styleArray2 = get_style(Alignment::HORIZONTAL_CENTER, Alignment::VERTICAL_CENTER);

$styleArray3 = get_style(Alignment::HORIZONTAL_RIGHT, Alignment::VERTICAL_CENTER);

$styleArray4 = get_style(Alignment::HORIZONTAL_LEFT, Alignment::VERTICAL_CENTER);

$styleArray5 = get_style(Alignment::HORIZONTAL_CENTER, Alignment::VERTICAL_CENTER, 1);

$styleArray6 = get_style(Alignment::HORIZONTAL_RIGHT, Alignment::VERTICAL_CENTER, 1);

// 新建styleArray
function get_style($horizontal, $vertical, $borders = null)
{
    $styleArray = [
        'alignment' => [
            'horizontal' => $horizontal,
            'vertical'   => $vertical,
        ],
    ];
    if (is_array($borders)) {
        $styleArray['borders'] = $borders;
    } else if ($borders == null) {

    } else {
        $styleArray['borders'] = [
            'top'    => [
                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
            ],
            'bottom' => [
                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
            ],
            'left'   => [
                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
            ],
            'right'  => [
                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
            ],
        ];
    }
    return $styleArray;
}

function isselect($value)
{
    if ($value == 'on') {
        $output = '■  ';
    } else {
        $output = '□  ';
    }
    return $output;
}

// 填充单元格
function fill_cell($style, $part, $cell, $val, $merge = null)
{
    global $spreadsheet;
    if ($merge != null) {
        $spreadsheet->getActiveSheet()->mergeCells($merge);
    }
    $spreadsheet->getActiveSheet()->getStyle($part)->getAlignment()->setWrapText(true);
    if ($style != null) {
        $spreadsheet->getActiveSheet()->getStyle($part)->applyFromArray($style);
    }

    $spreadsheet->getActiveSheet()->setCellValue($cell, $val);
}

// 设置横向纵向
function set_horizontal($isHorizontal = true, $isFit = true)
{
    global $spreadsheet;
    if ($isHorizontal) {
        $spreadsheet->getActiveSheet()->getPageSetup()
            ->setOrientation(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE);
    } else {
        $spreadsheet->getActiveSheet()->getPageSetup()
            ->setOrientation(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_PORTRAIT);
    }
    $spreadsheet->getActiveSheet()->getPageSetup()
        ->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);
    $spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage($isFit);
}

// 冻结窗格
function set_freeze($cell)
{
    global $spreadsheet;
    $spreadsheet->getActiveSheet()->freezePane($cell);
}

// 重复打印某一列
function set_repeatAtLeft($startAndEndColumn)
{
    global $spreadsheet;
    $spreadsheet->getActiveSheet()->getPageSetup()->setColumnsToRepeatAtLeft($startAndEndColumn);
}

// 设置适应宽度
function set_fitToWidth($w)
{
    global $spreadsheet;
    $spreadsheet->getActiveSheet()->getPageSetup()->setFitToWidth($w);
}

// 设置列宽度
function setColumnWidth($column, $w)
{
    global $spreadsheet;
    $spreadsheet->getActiveSheet()->getColumnDimension($column)->setWidth($w);
}

// 设置行高度
function setRowHeight($row, $h)
{
    global $spreadsheet;
    $spreadsheet->getActiveSheet()->getRowDimension($row)->setRowHeight($h);
}

// 设置字体粗细
function setBold($column, $isBold = true)
{
    global $spreadsheet;
    $spreadsheet->getActiveSheet()->getStyle($column)->getFont()->setBold($isBold);
}

// 填充图片
function fill_img($img, $cell, $w, $h)
{
    global $spreadsheet;
    if ($img == '') {
        $haveimg = false; //没有图片
    } else {
        $path     = $img;
        $pathinfo = pathinfo($path);
        if ($pathinfo["extension"] == 'pdf') {
            $img     = pdficon();
            $haveimg = true;
        } else {
            $haveimg = true;
        }
    }

    if ($haveimg) {
        preg_match('/.(jpg|gif|bmp|jpeg|png)/i', $img, $imgformat);
        $imgformat = $imgformat[1];
        switch ($imgformat) {
            case "jpg":
            case "jpeg":
                $img = imagecreatefromjpeg($img);
                break;
            case "bmp":
                $img = imagecreatefromwbmp($img);
                break;
            case "gif":
                $img = imagecreatefromgif($img);
                break;
            case "png":
                $img = imagecreatefrompng($img);
                break;
        }
        if ($img) {
            // Add a drawing to the worksheet
            $drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
            $drawing->setName('img');
            $drawing->setDescription('img');
            $drawing->setImageResource($img);
            $drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
            $drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
            $drawing->setResizeProportional(true);
            $drawing->setWidthAndHeight($w, $h); //设置图片最大宽度 高度
            // $drawing->setWidth($width);
            $drawing->setCoordinates($cell);
            $drawing->setOffsetX(5);
            $drawing->setOffsetY(5);
            $drawing->setWorksheet($spreadsheet->getActiveSheet());
        }
    }
}

// 加页
function add_sheet($index, $name)
{
    global $spreadsheet;
    $myWorkSheet = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($spreadsheet, $name);
    $spreadsheet->addSheet($myWorkSheet, $index);
}

// 输出
function set_writer($type = null, $outputName = null)
{
    // Set active sheet index to the first sheet, so Excel opens this as the first sheet
    global $spreadsheet;
    global $productall;
    global $cpsform;
    if ($type != null) {
        switch ($type) {
            case 'HA':
                $form_client = isset($productall['client']) ? $productall['client'] . "_" : "";
                $nt          = date("md", time());
                $outputName  = str_replace(",", "", $type . "_" . $form_client . $nt);
                break;
            case 'CPS':
                $nt         = date("md", time());
                $outputName = $type . "_" . $cpsform['client'] . "_" . $nt;
                break;
            default:
                $outputName = ($outputName != null && $outputName != "") ? $outputName : "untitled";
                break;
        }
    } else if ($type == null && $outputName != null && strlen($outputName) > 0) {
        $outputName = $outputName;
    } else {
        $outputName = ($outputName != null && $outputName != "") ? $outputName : "untitled";
    }

    $spreadsheet->setActiveSheetIndex(0);

    $output      = ($_GET['action'] == 'formdown') ? 1 : 0;
    $filenameout = $outputName . '.xlsx';

    if ($output) {
        // Redirect output to a client’s web browser (Xlsx)
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename=' . "$filenameout");
        header('Cache-Control: max-age=0');
        // If you're serving to IE 9, then the following may be needed
        header('Cache-Control: max-age=1');
        // If you're serving to IE over SSL, then the following may be needed
        header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
        header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
        header('Pragma: public'); // HTTP/1.0

        $writer = PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save('php://output');
    } else {
        $writer = new Xlsx($spreadsheet);
        $writer->save('../output/' . $filenameout);

        $FILEURL   = PRINTURL . $filenameout;
        $MSFILEURL = MSFILEURL . urlencode($FILEURL);

        Header("Location:{$MSFILEURL}");
    }
    exit;
}

// HA第3页
function set_ha_p3()
{
    global $spreadsheet;
    global $styleArray1;
    global $productall;

    $productp3 = $productall['productp3'];

    $spreadsheet->setActiveSheetIndex(2);
    fill_cell(null, 'E2', 'E2', $productp3['guest']);
    fill_cell(null, 'C2', 'C2', $productp3['styleno']);
    $startarr = 4;
    if ($productp3['a1']['formnum'] > 13) {
        $spreadsheet->getActiveSheet()->insertNewRowBefore(18, $productp3['a1']['formnum'] - 13);
    }
    for ($j = 1; $j < count($productp3['a1']); $j++) {
        $col    = chr(65 + $j);
        $carray = $productp3['a1']['c' . $j];
        foreach ($carray as $key => $item) {
            fill_cell($styleArray1, $col . ($key + $startarr), $col . ($key + $startarr), $item);
        }
    }
    set_horizontal(true);
}

// 格式化表格
function format_form($style, $row)
{
    global $spreadsheet;
    $spreadsheet->getActiveSheet()->getStyle('A' . $row)->applyFromArray($style);
    for ($i = 1; $i <= 19; $i++) {
        $col = chr(64 + $i);
        $spreadsheet->getActiveSheet()->getStyle($col . $row)->applyFromArray($style);
    }
}

// HA第4页
function set_ha_p4()
{
    global $spreadsheet;
    global $styleArray1;
    global $styleArray2;
    global $styleArray4;
    global $productall;

    $productp1 = $productall['productp1'];
    $productp4 = $productall['productp4'];

    $spreadsheet->setActiveSheetIndex(3);

    $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(15);
    $spreadsheet->getActiveSheet()->getStyle('A1')->getFont()->setSize(14);
    $spreadsheet->getActiveSheet()->getStyle('A1')->getFont()->setBold(true);
    for ($i = 1; $i <= 18; $i++) {
        $spreadsheet->getActiveSheet()->getColumnDimension(chr(65 + $i))->setWidth(8);
    }
    fill_cell($styleArray2, 'A1', 'A1', '生产通知单', 'A1:S1');
    fill_cell(null, 'A2', 'A2', '客户');
    fill_cell($styleArray2, 'B2', 'B2', $productp4['guest'], 'B2:F2');
    fill_cell(null, 'A3', 'A3', '开单日期');
    fill_cell($styleArray2, 'B3', 'B3', $productp1['alist']['a1'], 'B3:F3');
    fill_cell(null, 'G2', 'G2', '制单号');
    fill_cell($styleArray2, 'H2', 'H2', $productp1['jobno'], 'H2:M2');
    fill_cell(null, 'G3', 'G3', '款号');
    fill_cell($styleArray2, 'H3', 'H3', $productp4['styleno'], 'H3:M3');
    fill_cell(null, 'N2', 'N2', '生产部门');
    fill_cell($styleArray2, 'P2', 'P2', '', 'P2:S2');
    fill_cell(null, 'N3', 'N3', '落货日期');
    fill_cell($styleArray2, 'P3', 'P3', $productp1['alist']['a2'], 'P3:S3');

    if ($productp4['a1']['oarr']['o10'] == "0" || $productp4['a1']['oarr']['o10'] == "") {
        $row = 4;
        for ($i = 2; $i <= 18; $i++) {
            $col = chr(65 + $i);
            $spreadsheet->getActiveSheet()->getStyle($col . $row)->applyFromArray($styleArray1);
        }

        fill_cell($styleArray1, 'A4:B4', 'A4', '大货尺寸表', 'A4:B4');
        fill_cell($styleArray1, 'F4', 'F4', '缩水率');
        fill_cell($styleArray1, 'H4', 'H4', '直:');
        fill_cell($styleArray1, 'L4', 'L4', '横:');
        fill_cell($styleArray2, 'I4', 'I4', $productp4['a1']['oarr']['o1'], 'I4:K4');
        fill_cell($styleArray2, 'M4', 'M4', $productp4['a1']['oarr']['o0'], 'M4:O4');
        fill_cell($styleArray1, 'R4', 'R4', '公差 +/-');
        fill_cell($styleArray1, 'S4', 'S4', '缩水率');

        $row = 5;
        format_form($styleArray1, $row);
        for ($i = 1; $i <= 11; $i++) {
            if ($i == 1) {
                $A5 = $productp4['a1']['oarr']['o12'] == "1" ? "IN" : "CM";
                fill_cell($styleArray1, 'A' . $row, 'A' . $row, $A5);
            } else if ($i == 10) {
                fill_cell($styleArray1, 'R' . $row, 'R' . $row, '');
            } else if ($i == 11) {
                fill_cell($styleArray1, 'S' . $row, 'S' . $row, '');
            } else {
                $col     = chr(64 + $i * 2 - 2);
                $colNext = chr(64 + $i * 2 - 1);
                fill_cell($styleArray1, $col . $row . ":" . $colNext . $row, $col . $row, $productp4['a1']['oarr']['o' . $i], $col . $row . ":" . $colNext . $row);
            }
        }

        $row         = 6;
        $insertCount = count($productp4['a1']['carr']['c1']);
        for ($i = 1; $i <= $productp4['a1']['carr']['colnum']; $i++) {
            if ($i == 1) {
                for ($j = 0; $j < $insertCount; $j++) {
                    format_form($styleArray1, 5 + $j);
                }
            }
            foreach ($productp4['a1']['carr']['c' . $i] as $index => $item) {
                if ($i == 1) {
                    fill_cell($styleArray1, 'A' . ($row + $index), 'A' . ($row + $index), $item);
                } else if ($i == 10) {
                    fill_cell($styleArray1, 'R' . ($row + $index), 'R' . ($row + $index), $item);
                } else if ($i == 11) {
                    fill_cell($styleArray1, 'S' . ($row + $index), 'S' . ($row + $index), $item);
                } else {
                    $col     = chr(64 + $i * 2 - 2);
                    $colNext = chr(64 + $i * 2 - 1);
                    fill_cell($styleArray1, $col . ($row + $index) . ":" . $colNext . ($row + $index), $col . ($row + $index), $item, $col . ($row + $index) . ":" . $colNext . ($row + $index));
                }
            }
        }
        // remark
        $row = $row + $insertCount + 2;
        for ($i = 0; $i < count($productp4['a1']['remarkarr']['e1']); $i++) {
            fill_cell($styleArray4, 'B' . ($row + $i), 'B' . ($row + $i), $productp4['a1']['remarkarr']['e1'][$i], 'B' . ($row + $i) . ":" . 'Q' . ($row + $i));
        }
        set_horizontal(true);

    } else {
        fill_img($productp4['a1']['oarr']['o11'], 'A5', 1400, 750);
        set_horizontal(true); //打印橫向 A4

        // page 5
        add_sheet(4, "remark");
        $spreadsheet->setActiveSheetIndex(4);
        $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('A1')->getFont()->setSize(14);
        $spreadsheet->getActiveSheet()->getStyle('A1')->getFont()->setBold(true);
        for ($i = 1; $i <= 18; $i++) {
            $spreadsheet->getActiveSheet()->getColumnDimension(chr(65 + $i))->setWidth(8);
        }

        fill_cell($styleArray2, 'A1', 'A1', '生产通知单', 'A1:S1');
        fill_cell(null, 'A2', 'A2', '客户');
        fill_cell($styleArray2, 'B2', 'B2', $productp4['guest'], 'B2:F2');
        fill_cell(null, 'A3', 'A3', '开单日期');
        fill_cell($styleArray2, 'B3', 'B3', $productp1['alist']['a1'], 'B3:F3');
        fill_cell(null, 'G2', 'G2', '制单号');
        fill_cell($styleArray2, 'H2', 'H2', $productp1['jobno'], 'H2:M2');
        fill_cell(null, 'G3', 'G3', '款号');
        fill_cell($styleArray2, 'H3', 'H3', $productp4['styleno'], 'H3:M3');
        fill_cell(null, 'N2', 'N2', '生产部门');
        fill_cell($styleArray2, 'P2', 'P2', '', 'P2:S2');
        fill_cell(null, 'N3', 'N3', '落货日期');
        fill_cell($styleArray2, 'P3', 'P3', $productp1['alist']['a2'], 'P3:S3');

        // remark
        $row = 5;
        for ($i = 0; $i < count($productp4['a1']['remarkarr']['e1']); $i++) {
            fill_cell($styleArray4, 'B' . ($row + $i), 'B' . ($row + $i), $productp4['a1']['remarkarr']['e1'][$i], 'B' . ($row + $i) . ":" . 'Q' . ($row + $i));
        }
        set_horizontal(true); //打印橫向 A4
    }
}

function getforexcate($forex)
{
    switch ($forex) {
        case 1:
            $output = 'USD';
            break;
        case 2:
            $output = 'HKD';
            break;
        case 3:
            $output = 'RMB';
            break;
        case 4:
            $output = 'EUR';
            break;
        case 5:
            $output = 'JPY';
            break;

        default:
            $output = 'USD';
            break;
    }
    return $output;
}

// PCS页面打印适应
function set_print_pcs($freeze)
{
    set_freeze($freeze);
    // 重复左侧
    set_repeatAtLeft(array('A', 'A'));
    // 横向不缩放
    set_horizontal(true, false);
    // 所有行打印在一页
    set_fitToWidth(0);
}
