<?php

error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);

date_default_timezone_set('Europe/Paris');

require_once("Classes/PHPExcel.php");
/** Include PHPExcel_IOFactory */
require_once dirname(__FILE__) . '/Classes/PHPExcel/IOFactory.php';

function endswith($string, $test) {
    $strlen = strlen($string);
    $testlen = strlen($test);
    if ($testlen > $strlen) return false;
    return substr_compare($string, $test, $strlen - $testlen, $testlen) === 0;
}

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
	if($cell_info->dataType == "inlineStr")
	{
		$cell_info->value = $cell_info->value->getPlainText();
		//var_dump($cell_info->value);
	}
	
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
			//printf("Stopped scanning at two subsequent blank cells\n");
			break;
		}
		if($cell_info->value != NULL)
		{
			if($cell_info->bold)
			{
				//printf("Found new section %s\n", $cell_info->value);
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

function printUsage()
{
	printf("Usage:\n\tconvert <filename> (JSON or XLSX file)\n\nif given a JSON file it converts to XLSX (Excel2007) and the other way around\n\n");
}

function convertJSONToExcel($src_filename, $dst_filename)
{
	$workbook = new PHPExcel();
	$sheet = $workbook->getSheet();

	$jsondata = json_decode(file_get_contents($src_filename), true);
	//print_r($jsondata);
	$x = 1;
	foreach($jsondata as $locale => $lang)
	{
		printf("Converting language %s...\n", $locale);
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
	$objWriter->save($dst_filename);
}

function convertExcelToJSON($src_filename, $dst_filename)
{
	$workbook = PHPExcel_IOFactory::load($src_filename);
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
		$fname = $dst_filename . "_" . $name . ".json";
		printf("Writing language file %s...\n", $fname);
		file_put_contents($fname, $json);
	}
}

printf("ExcelStack convert 1.0 - by Bison Montana the Lord of Salmon\n\n");

if($argc != 2)
{
	printUsage();
	exit(0);
}
$filename = $argv[1];
if(!file_exists($filename))
{
	printf("%s file does not exist\n", $filename);
	exit(1);
}
if(endswith(strtolower($filename), ".json"))
{
	$dst_filename = substr_replace($filename, "xlsx", -4);
	printf("Converting JSON file $filename to excel format $dst_filename\n");
	convertJSONToExcel($filename, $dst_filename);
}
else if(endswith(strtolower($filename), ".xlsx"))
{
	printf("Converting Excel file $filename to json format\n");
	$dst_filename = substr_replace($filename, "", -5);
	convertExcelToJSON($filename, $dst_filename);
}
else
{
	printf("File format not recognized (suffix must be either .json or .xlsx)\n");
}

