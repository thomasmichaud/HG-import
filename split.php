<?php

/** Error reporting */
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/Paris');

define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

/** Include PHPExcel */
require_once dirname(__FILE__) . '/Classes/PHPExcel.php';

$file = "import.xlsx";
$loadSheets = array('MAREE - Tableau 1', 'MAREE - Tableau 1-2', 'Viande - Tableau 1', 'Volaille - Tableau 1', 'F&L - Tableau 1', 'Epicerie Cremerie - Tableau 1');

foreach($loadSheets as $sheet){

	$objReader = new PHPExcel_Reader_Excel2007();

	//if we dont need any formatting on the data
	$objReader->setReadDataOnly();

	//load only certain sheets from the file

	$objReader->setLoadSheetsOnly($sheet);

	$objPHPExcel = $objReader->load($file);

	$outputfile = "split-".str_replace(" ","",$sheet).".xlsx";

	$objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
	$objWriter->save($outputfile);

}



?>