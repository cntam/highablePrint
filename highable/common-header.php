<?php
session_start();
require_once 'autoloadconfig.php'; //判断是否在线
require_once 'img.php';

//$spreadsheet = new Spreadsheet();
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('../template/' . $xlsxName . '.xlsx');
// $sheet       = $spreadsheet->getActiveSheet();
$spreadsheet->setActiveSheetIndex(0);
$spreadsheet->getDefaultStyle()->getFont()->setName('Microsoft Yahei');
$spreadsheet->getDefaultStyle()->getFont()->setSize(12);

$styleArray = [

    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
        'vertical'   => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
    ],

    'borders'   => [
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
    ],

];

$styleArraytop = [

    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
        'vertical'   => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP,
    ],

    'borders'   => [
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
    ],

];

$styleArray1 = [
    'alignment' => [
        'horizontal'  => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
        'vertical'    => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP,
        'wrapText'    => true,
        'ShrinkToFit' => true,
    ],
    'font'      => [
        'Size' => '10',
    ],

    'borders'   => [
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

    ],

];

$styleArray2 = [
    'alignment' => [
        'horizontal'  => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
        'vertical'    => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
        'wrapText'    => true,
        'ShrinkToFit' => true,
    ],
    'font'      => [
        'Size' => '10',
    ],
];

$styleArray3 = [
    'alignment' => [
        'horizontal'  => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT,
        'vertical'    => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
        'wrapText'    => true,
        'ShrinkToFit' => true,
    ],
    'font'      => [
        'Size' => '10',
    ],
];

$styleArray4 = [
    'alignment' => [
        'horizontal'  => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
        'vertical'    => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
        'wrapText'    => true,
        'ShrinkToFit' => true,
    ],
    'font'      => [
        'Size' => '10',
    ],
];

function isselect($value)
{
    if ($value == 'on') {
        $output = '■  ';
    } else {
        $output = '□  ';
    }
    return $output;
}

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

function set_horizontal($isHorizontal = true)
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
    $spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true);
}

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
        $width  = imagesx($img);
        $height = imagesy($img);

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
        // return $drawing->getHeight();
    }
}

function add_sheet($index, $name)
{
    global $spreadsheet;
    $myWorkSheet = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($spreadsheet, $name);
    $spreadsheet->addSheet($myWorkSheet, $index);
}

function set_writer()
{

    // Set active sheet index to the first sheet, so Excel opens this as the first sheet
    global $spreadsheet;
    global $productall;

//    $form_type    = isset($productall['type']) ? $productall['type'] . "_" : "";
//    $form_client  = isset($productall['client']) ? $productall['client'] . "_" : "";
//    $modification = isset($productall['modification']) ? date("YmdHis", $productall['modification']) : "";
//    $outputName   = str_replace(",", "", $form_type . $form_client . $modification);
    $outputName = 'qty';
    $spreadsheet->setActiveSheetIndex(0);

    $output = ($_GET['action'] == 'formdown') ? 1 : 0;
    // $nt          = date("YmdHis", time()); //转换为日期。
    // $filenameout = $outputName . $nt . '.xlsx';
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

        // $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
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
function set_p3()
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
function format_form($style, $row)
{
    global $spreadsheet;
    $spreadsheet->getActiveSheet()->getStyle('A' . $row)->applyFromArray($style);
    for ($i = 1; $i <= 19; $i++) {
        $col = chr(64 + $i);
        $spreadsheet->getActiveSheet()->getStyle($col . $row)->applyFromArray($style);
    }
}
function set_p4()
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
