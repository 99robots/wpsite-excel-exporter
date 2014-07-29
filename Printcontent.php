<?php
require_once('ExportHeading.php');
$letter_range=range('A','Z');
function print_content($post_content,$noofrows) {
	global $objPHPExcel;
	global $letter_range; 	
	global $objWriter;	
	global $objRichText;
	$counter=0;
	$cell_name="";
	foreach($post_content as $content) {
		$cell_name = $letter_range[$counter].$noofrows;
		$counter++;
		$value = $content;
		$objPHPExcel->getActiveSheet()->SetCellValue($cell_name, $value);
	}
	
}
?>