<?php
require_once 'aidenfunc.php';
$cpsform =  $_SESSION['SampleChart'];

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet = new Spreadsheet();

//$sheet = $spreadsheet->getActiveSheet();
//$spreadsheet->getActiveSheet()->setTitle("CPS");
//$sheet->setCellValue('A1', 'Hello World !');
$spreadsheet->getDefaultStyle()->getFont()->setName('微软雅黑');
$spreadsheet->getDefaultStyle()->getFont()->setSize(12);
//$spreadsheet->getActiveSheet()->getDefaultRowDimension()->setRowHeight(50); //行默认高度
//$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(20);  //列宽度

$border = \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN;
$h_center = \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER;
$v_center = \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER;
$fill_solid = \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID;
$v_top = \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP;
$styleArray1 = [
    'alignment' => [
        'vertical' => $v_center,
    ],
    'borders' => [
        'top' => [
            'borderStyle' => $border,
        ],
        'bottom' => [
            'borderStyle' => $border,
        ],
        'left' => [
            'borderStyle' => $border,
        ],
        'right' => [
            'borderStyle' => $border,
        ],
    ],

];
$styleArraycenter = [
    'alignment' => [
        'vertical' => $v_center,
        'horizontal' => $h_center,
        'wrapText' => true,
        'ShrinkToFit'=>true,
    ],
    'borders' => [
        'top' => [
            'borderStyle' => $border,
        ],
        'bottom' => [
            'borderStyle' => $border,
        ],
        'left' => [
            'borderStyle' => $border,
        ],
        'right' => [
            'borderStyle' => $border,
        ],
    ],

];
$styleArray = [
    'alignment' => [
        'vertical' => $v_center,
    ],
    'borders' => [
        'top' => [
            'borderStyle' => $border,
        ],
        'bottom' => [
            'borderStyle' => $border,
        ],
        'left' => [
            'borderStyle' => $border,
        ],
        'right' => [
            'borderStyle' => $border,
        ],
    ],

];
$styleArrayborder = [
    'borders' => [
        'top' => [
            'borderStyle' => $border,
        ],
        'bottom' => [
            'borderStyle' => $border,
        ],
        'left' => [
            'borderStyle' => $border,
        ],
        'right' => [
            'borderStyle' => $border,
        ],
    ]
];
$styleArraylefttop = [
    'alignment' => [
        'vertical' => $v_top ,
    ],
    'borders' => [
        'top' => [
            'borderStyle' => $border,
        ],
        'bottom' => [
            'borderStyle' => $border,
        ],
        'left' => [
            'borderStyle' => $border,
        ],
        'right' => [
            'borderStyle' => $border,
        ],
    ]
];
function getforexcate($forex) {
    switch ($forex){
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
function isselect($value){
    if ( $value == 'on') {
        $output = '■  ';
    } else {
        $output = '□  ';
    }
    return $output;
}
function add_sheet_header($sheet){
    global $styleArrayborder ;
    global $styleArraylefttop;
    for($i=1;$i<11;$i++){
        $sheet->getStyle("A".$i)->applyFromArray($styleArrayborder);
    }
    $sheet->setCellValue("A1", 'Designer：');
    $sheet->setCellValue("A2", 'Sample order no.：');
    $sheet->setCellValue("A3", 'Sketch：');
    $sheet->getStyle("A3")->applyFromArray($styleArraylefttop);
    $sheet->setCellValue("A4", 'Collection Name:');
    $sheet->setCellValue("A5", 'Style No.:');
    $sheet->setCellValue("A6", 'Main Fabric:');
    $sheet->getStyle("A6")->applyFromArray($styleArraylefttop);
    $sheet->setCellValue("A7", 'Lining:');
    $sheet->setCellValue("A8", 'Trim Fabric:');
    $sheet->setCellValue("A9", 'Trims:');
    $sheet->setCellValue("A10", 'Remarks:');
    $sheet->getStyle("A10")->applyFromArray($styleArraylefttop);

    $sheet->getRowDimension('1')->setRowHeight(20); //列高度
    $sheet->getRowDimension('2')->setRowHeight(20); //列高度
    $sheet->getRowDimension('3')->setRowHeight(160); //列高度
    $sheet->getRowDimension('8')->setRowHeight(160); //列高度
    $sheet->getRowDimension('10')->setRowHeight(160); //列高度

    $sheet->getColumnDimension('A')->setWidth(20);  //列宽度
    $sheet->getColumnDimension('B')->setWidth(30);  //列宽度
    $sheet->getColumnDimension('C')->setWidth(30);  //列宽度
    $sheet->getColumnDimension('D')->setWidth(30);  //列宽度
    $sheet->getColumnDimension('E')->setWidth(30);  //列宽度
    $sheet->getColumnDimension('F')->setWidth(30);  //列宽度


}
function add_sheet_data($sheet,$col,$item){
    global $styleArraycenter;
    fill_img($sheet,$item['a3'],$col.'3',200,200);
    $sheet->setCellValue($col.'1', $item['a1']);
    $sheet->setCellValue($col.'2', $item['a2']);
    $sheet->setCellValue($col.'4', $item['a4']);
    $sheet->setCellValue($col.'5', $item['a5']);
    $sheet->setCellValue($col.'6', stripcslashes($item['a6']));
    $sheet->setCellValue($col.'7', stripcslashes($item['a7']));
    $sheet->setCellValue($col.'8', stripcslashes($item['a8']));
    $sheet->setCellValue($col.'9', stripcslashes($item['a9']));
    $sheet->setCellValue($col.'10',stripcslashes($item['a10']));
    for($i=1;$i<11;$i++){
        $sheet->getStyle($col.$i)->applyFromArray($styleArraycenter);
    }
}
function fill_img($sheet, $img, $cell, $w, $h)
{
    global $styleArraycenter;
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
        $drawing->setWorksheet($sheet);
        $sheet->getStyle($cell)->applyFromArray($styleArraycenter);
    }
}
$sheet_data_array = array_chunk($cpsform['info'],5);
foreach ($sheet_data_array as $index => $item){
    $myWorkSheet = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($spreadsheet, 'sheet'.($index+1));
    $spreadsheet->addSheet($myWorkSheet, $index);
    $sheet = $spreadsheet->getSheet($index);
    add_sheet_header($sheet);
    $col_index = array('B','C','D','E','F');
    if(count($item)>0) {
        for ($i = 0; $i < count($item); $i++) {
            add_sheet_data($sheet, $col_index[$i], $item[$i]);
        }
    }
    $spreadsheet->getActiveSheet()->getPageSetup()
        ->setOrientation(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE); //打印橫向
    $spreadsheet->getActiveSheet()->getPageSetup()
        ->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);//打印橫向 A4
    $spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true); //将工作表调整为一页
}



// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$spreadsheet->setActiveSheetIndex(0);
unset($_SESSION['samplechart'] ); //注销SESSION

$filenameout = 'SampleChart_';

outExcel($spreadsheet,$filenameout);
