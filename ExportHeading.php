<?php
require_once('PHPExcel/IOFactory.php');
require_once('PHPExcel/Writer/Excel5.php');


$objPHPExcel = new PHPExcel();
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objRichText = new PHPExcel_RichText();
// Set properties
$objPHPExcel->getProperties()->setCreator("author");
$objPHPExcel->getProperties()->setLastModifiedBy("author");
$objPHPExcel->getProperties()->setTitle("communityanswers");
$objPHPExcel->getProperties()->setSubject("Office 2003Excel");
$objPHPExcel->getProperties()->setDescription("Office 2003Excel");


// Add some data
$objPHPExcel->setActiveSheetIndex(0);
for ($i = 0;$i <= 24; $i++) {
	$objPHPExcel->getActiveSheet()->getColumnDimension(chr(65+$i))->setAutoSize(true);
}
$objPHPExcel->getActiveSheet()->getStyle('Y1')->getAlignment()->setWrapText(true); 

$objPHPExcel->getActiveSheet()->setCellValue('J1', 'Physical Address:')
					  ->setCellValue('T1','Contact Information:');

$letters = range('A','Z');



function ExportToExcel($tittles,$excel_name) {
	global $objPHPExcel;
	global $letters; 	
	global $objWriter;	
	global $objRichText;
		$countdown =0;
		$c=1;
		$cell_name="";
		foreach($tittles as $tittle) {
			$cell_name = $letters[$countdown]."2";
			$countdown++;
			$next_cell=$letters[$c]."2";
			$c++;
			$value = $tittle;
			$objPHPExcel->getActiveSheet()->SetCellValue($cell_name, $value);
			// Make bold cells
			//$objPHPExcel->getActiveSheet()->getStyle("$cell_name:$next_cell")->getFont()->setBold(true);
			$styleArray = array( 'font' => array( 'bold' => true, 'underline' => PHPExcel_Style_Font::UNDERLINE_SINGLE),);
			$objPHPExcel->getActiveSheet()->getStyle("$cell_name:$next_cell")->applyFromArray($styleArray);
		}	

}
?>