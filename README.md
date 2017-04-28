# ExcelStack

This CLI tool convert exported NStack language translations to Excel (XLSX) format. The generated excel workbook can then be edited and exported to a JSON format which is importable by the NStack webinterface.

ExcelStack is written as a PHP script utilizing the PHP Excel library (included). The main script convert.php is supposed to be run from the command line and requires a valid php installation with the PHP interpreter added to the system PATH variable (PHP installer should take care of this).

## Usage

	php convert.php <filename> (JSON or XLSX file)

If given a JSON file it converts to XLSX (Excel2007) ready for human editing.
Same file can then be converted back to a set of JSON files (one for each language), which can then be imported into NStack.
