<?php

/** Error reporting */
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/Paris');

define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

/** Include PHPExcel */
require_once dirname(__FILE__) . '/Classes/PHPExcel.php';
require_once dirname(__FILE__) . '/merge.php';

//Données de traitement
//$file = "test.xlsx";
$file = $_GET["source"].".xlsx";
$objReader = new PHPExcel_Reader_Excel2007();
$worksheetData = $objReader->listWorksheetInfo($file);
$totalRows     = $worksheetData[0]['totalRows'];
$totalColumns  = $worksheetData[0]['totalColumns'];
$coloneSplit = $totalColumns / 2;

//Données
$alphabet = range('A', 'Z');

//On créé le constructeur du filtre
class MyReadFilter implements PHPExcel_Reader_IReadFilter
{
	private $_startRow = 0;

	private $_endRow = 0;

	private $_columns = array();

	public function __construct($startRow, $endRow, $columns) {
		$this->_startRow	= $startRow;
		$this->_endRow		= $endRow;
		$this->_columns		= $columns;
	}

	public function readCell($column, $row, $worksheetName = '') {
		if ($row >= $this->_startRow && $row <= $this->_endRow) {
			if (in_array($column,$this->_columns)) {
				return true;
			}
		}
		return false;
	}
}


//On récupère la première colonne
$objReader->setReadFilter( new MyReadFilter(1,$totalRows ,range('A',$alphabet[$coloneSplit-1])) );
$objPHPExcel = $objReader->load($file);
$objPHPExcel->getActiveSheet()->setTitle("S1");


//Export premiere colonne
$objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
$objWriter->save("export-c1.xlsx");


//On récupère la deuxieme colonne
$objReader2 = new PHPExcel_Reader_Excel2007();
$objReader2->setReadFilter( new MyReadFilter(1,$totalRows ,range($alphabet[$coloneSplit], 'J')));
$objPHPExcel2 = $objReader2->load($file);
$objPHPExcel2->getActiveSheet()->setTitle("S2");

//Export de la deuxime colonne
$objWriter2 = new PHPExcel_Writer_Excel2007($objPHPExcel2);
$objWriter2->save("export-c2.xlsx");

//On merge les fichiers temporaires créés
$outputFile = mergeSheet('export-c1.xlsx', 'export-c2.xlsx', $coloneSplit, $alphabet, $totalRows, $_GET["source"]);

//Suppression des fichiers temporaires
unlink('export-c1.xlsx');
unlink('export-c2.xlsx');

//On transforme le fichier exporté en CVS
$objReaderExport = new PHPExcel_Reader_Excel2007();
$objPHPExcelExport = $objReaderExport->load($outputFile);
$objWriterExport = new PHPExcel_Writer_CSV($objPHPExcelExport);
$objWriterExport->setUseBOM(true);
$objWriterExport->setDelimiter(';');
$objWriterExport->setEnclosure('');
//$objWriter->setLineEnding("\r\n"); 
$objWriter->save("export.csv");

//On supprime le fichier d'export xlsx
//unlink($outputFile);


/*
$objWriter = new PHPExcel_Writer_CSV($objPHPExcel);
$objWriter->setUseBOM(true);
$objWriter->setDelimiter(';');
$objWriter->setEnclosure('');
//$objWriter->setLineEnding("\r\n"); 
$objWriter->save("export.csv");
*/

?>