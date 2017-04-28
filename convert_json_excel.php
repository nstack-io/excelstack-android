<?php
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);

date_default_timezone_set('Europe/Paris');

require_once("Classes/PHPExcel.php");
/** Include PHPExcel_IOFactory */
require_once dirname(__FILE__) . '/Classes/PHPExcel/IOFactory.php';

function buildLanguageArray(&$sheet)
{
	$lang = array();

	$x = 1;
	$y = 1;
	while(true)
	{
		$value = $sheet->getCellByColumnAndRow($x, $y)->getValue();
		if($value == NULL)
			break;
		//var_dump($value);
		$dataType = PHPExcel_Cell_DataType::dataTypeForValue($value);
		//var_dump($dataType);
		$lang[$value] = array('x' => $x);
		$x++;
	}
	return $lang;
}

function getCellInfo(&$sheet, $x, $y)
{
	$style = $sheet->getStyleByColumnAndRow($x, $y);
	$cell_info = new stdClass();
	$cell_info->value = $sheet->getCellByColumnAndRow($x, $y)->getValue();
	$cell_info->bold = $style->getFont()->getBold();
	$cell_info->color = $style->getFont()->getColor()->getRGB();
	$cell_info->dataType = PHPExcel_Cell_DataType::dataTypeForValue($cell_info->value);
	return $cell_info;
}

function buildSectionKeysArray(&$sheet, &$languages)
{
	$keys = array();
	$x = 0;
	$y = 3;
	$count = 0;

	$section_name = NULL;
	while(true)
	{
		$cell_info = getCellInfo($sheet, $x, $y);
		if($cell_info->value == NULL && $sheet->getCellByColumnAndRow($x, $y+1)->getValue() == NULL)
		{
			printf("Stopped scanning at two subsequent blank cells\n");
			break;
		}
		if($cell_info->value != NULL)
		{
			if($cell_info->bold)
			{
				printf("Found new section %s\n", $cell_info->value);
				$section_name = $cell_info->value;
			}
			else if($section_name != NULL)
			{
				//printf("\tKey: %s\n", $cell_info->value);
				$keys[$section_name][] = array('key' => $cell_info->value, 'row' => $y);
			}
			//var_dump($cell_info);
			$count++;
		}
		$y++;
	}
	printf("Scanned %d rows\n", $count);
	return $keys;
}

printf("ExcelStack convert 1.0 - by Bison Montana the Lord of Salmon\n\n");

//PHPExcel_Shared_Font::setAutoSizeMethod(PHPExcel_Shared_Font::AUTOSIZE_METHOD_EXACT);
$workbook = new PHPExcel();
$sheet = $workbook->getSheet();

$jsondata = json_decode(file_get_contents("eboks_nstack_export.json"), true);
//print_r($jsondata);
$x = 1;
foreach($jsondata as $locale => $lang)
{
	printf("%s\n", $locale);
	$y = 1;
	$sheet->getCellByColumnAndRow($x, $y)->setValue($locale);
	foreach($lang as $section_name => $section)
	{
		$y++;
		$sheet->getCellByColumnAndRow(0, $y)->setValue($section_name);
		$sheet->getStyleByColumnAndRow(0, $y)->getFont()->setBold(true);

		$y++;
		foreach($section as $key => $value)
		{
			$y++;
			$sheet->getCellByColumnAndRow($x, $y)->setValue($value);
			//if(getCellInfo($sheet, 1, $y)->value == NULL)
			$sheet->getCellByColumnAndRow(0, $y)->setValue($key);
		}
		$y++;
		
	}
	$sheet->getColumnDimension($sheet->getCellByColumnAndRow($x, 1)->getColumn())->setAutoSize(true);
	$x++;
}

$sheet->freezePaneByColumnAndRow(count($jsondata)+1, 2);

$objWriter = PHPExcel_IOFactory::createWriter($workbook, 'Excel2007');
$objWriter->save('eboks_translations.xlsx');
