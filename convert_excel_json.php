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

$workbook = PHPExcel_IOFactory::load("eboks.mobile.resources.nodes.xlsx");

$sheet = $workbook->getSheet();

printf("Sheet name: %s\n", $sheet->getTitle());

$languages = buildLanguageArray($sheet);
//print_r($languages);
$keys = buildSectionKeysArray($sheet, $languages);


$lang_json = array();
foreach ($languages as $name => $lang) 
{
	printf("Processing %s...\n", $name);
	$lang_json[$name] = array();
	foreach($keys as $section_name => $section)
	{
		$lang_json[$name][$section_name] = array();
		foreach($section as $val)
		{
			$ci = getCellInfo($sheet, $lang['x'], $val['row']);
			$lang_json[$name][$section_name][$val['key']] = $ci->value;
		}
	}
}

foreach ($lang_json as $name => $lang) 
{
	$json = json_encode($lang, JSON_PRETTY_PRINT | JSON_FORCE_OBJECT);
	file_put_contents("lang_" . $name . ".json", $json);
}
//print_r($lang_json);
