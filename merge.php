<?php


/** Include PHPExcel */
require_once dirname(__FILE__) . '/Classes/PHPExcel.php';

function mergeSheet($file1, $file2, $coloneSplit, $alphabet, $highRow, $outputname) {


//$filenames = array('doc1.xlsx', 'doc2.xlsx');

$objPHPExcel1 = PHPExcel_IOFactory::load($file1);
$objPHPExcel2 = PHPExcel_IOFactory::load($file2);

$objPHPExcel1->getActiveSheet()->fromArray(
    $objPHPExcel2->getActiveSheet()->rangeToArray($alphabet[$coloneSplit].'1:'.$alphabet[$coloneSplit*2 - 1].$highRow),
    null,
    'A' . ($objPHPExcel1->getActiveSheet()->getHighestRow() + 1)
);

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel1, 'Excel2007');
$objWriter->save($outputname.'-export.xlsx');

return $outputname.'-export.xlsx';
}

?>