<?php
// Author: Lucas Wasp, email: lucas@lucaswasp.com

set_time_limit(650);

require_once 'classes/PHPExcel.php';
require_once 'config.php';

$objPHPExcel = new PHPExcel(); 
$objPHPExcel->setActiveSheetIndex(0); 

$link = mysqli_connect ($host, $user, $password) or die('Could not connect: ' . mysqli_error());
mysqli_select_db($link, $db_name) or die('Could not select database');
$select = "SELECT * FROM `".$tbl_name."`";

mysqli_query($link, 'SET NAMES utf8;');
$export = mysqli_query($link, $select); 

$fields = mysqli_num_fields($export);



for ($i = 0, $letter = "A"; $i < $fields; $i++, $letter++) {
	$value = mysqli_fetch_field_direct($export, $i);
	$objPHPExcel->getActiveSheet()->SetCellValue($letter . '1', ($value -> name)); 
}
$i = 2;
$letter = "A";
while($row = mysqli_fetch_row($export)) {
	foreach($row as $value) {
		$objPHPExcel->getActiveSheet()->SetCellValue($letter++.$i, $value); 
	}
	$i++;
	$letter = "A";
}

$objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel); 
$objWriter->save("output/" . $file_name ); 

exit;
?>