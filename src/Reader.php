<?php

namespace Meridius\PhpExcel;

use Meridius\PhpExcel\ExcelFieldType;
use Meridius\PhpExcel\PhpExcelException;
use Meridius\Helpers\StringHelper;

class Reader extends \Nette\Object implements \Iterator {

	/** @var string */
	private $file;

	/** @var \PHPExcel */
	private $excel;

	/** @var \PHPExcel_Worksheet */
	private $activeSheet;

	/** @var string */
	private $excelColumnsRange = [];

	/** @var int */
	private $currentRow = 1;

	/** @var int */
	private $firstDataRow = 1;

	/** @var int */
	private $headerRow = 1;

	/** @var int */
	private $sheetIndex = 0;

	/** @var string|null */
	private $sheetName = null;

	/** @var array */
	private $sheetsToLoad = [];

	/** @var array */
	private $fieldsToRead = [];

	/** @var array */
	private $loweredFields = [];

	/** @var array */
	private $loweredToOriginalKeysMap = [];

	/** @var bool */
	private $deleteFileOnFinish = false;

	/** @var string[] */
	private $loweredFieldNameToExcelColumnMap = [];

	/**
	 *
	 * @param string $file
	 */
	public function __construct($file) {
		$this->file = $file;
	}

	public function __destruct() {
		if ($this->deleteFileOnFinish) {
			unlink($this->file);
		}
	}

	/**
	 * Opens excel file
	 * @return Reader
	 * @throws PhpExcelException
	 */
	public function open() {
		$reader = \PHPExcel_IOFactory::createReaderForFile($this->file);
		if (count($this->sheetsToLoad) > 0) {
			$reader->setLoadSheetsOnly($this->sheetsToLoad);
		}
		$reader->setReadDataOnly(true);
		$this->excel = $reader->load($this->file);
		if (is_null($this->sheetName)) {
			try {
				$this->activeSheet = $this->excel->getSheet($this->sheetIndex);
			} catch (\Exception $e) {
				throw new PhpExcelException("Sheet index '$this->sheetIndex' is out of range of excel indexes.");
			}
		} else {
			$this->activeSheet = $this->excel->getSheetByName($this->sheetName);
		}
		if (!($this->activeSheet instanceof \PHPExcel_Worksheet)) {
			throw new PhpExcelException("Sheet with name '$this->sheetName' does not exist in input excel.");
		}
		$this->mapHeaders();
		return $this;
	}

	/**
	 * @return \PHPExcel_Worksheet
	 */
	public function getActiveSheet() {
		return $this->activeSheet;
	}

	/**
	 * Sets fields to read (header cells)
	 * @param array $fields
	 * @return Reader
	 */
	public function setFieldsToRead(array $fields) {
		$this->fieldsToRead = $fields;
		foreach ($this->fieldsToRead as $key => $value) {
			$loweredKey = $this->lowerHeaderCellText($key);
			$this->loweredFields[$loweredKey] = $value;
			$this->loweredToOriginalKeysMap[$loweredKey] = $key;
		}
		return $this;
	}

	/**
	 * Sets mx range of colunmns to read
	 * @param string $from
	 * @param string $to
	 * @return Reader
	 */
	public function setExcelColumnsRange($from = 'A', $to = 'Z') {
		$this->excelColumnsRange = [];
		$columnIndex = $from;
		$to++;
		while ($columnIndex != $to) {
			$this->excelColumnsRange[] = $columnIndex;
			$columnIndex++;
		}
		return $this;
	}

	/**
	 * Sets sheet index to be read
	 * @param int $sheetIndex zero indexed
	 * @return Reader
	 */
	public function setSheetToRead($sheetIndex) {
		$this->sheetIndex = $sheetIndex;
		return $this;
	}

	/**
	 * Sets sheet name to be read
	 * @param string $sheetName
	 * @return Reader
	 */
	public function setSheetNameToRead($sheetName) {
		$this->sheetName = $sheetName;
		return $this;
	}

	/**
	 * Sets header row number
	 * @param int $headerRow
	 * @return Reader
	 */
	public function setHeaderRowNumber($headerRow) {
		$this->headerRow = $headerRow;
		return $this;
	}

	/**
	 * Sets sheet names to be load
	 * @param array $sheets
	 * @return Reader
	 */
	public function setSheetsToLoad(array $sheets) {
		$this->sheetsToLoad = $sheets;
		return $this;
	}

	/**
	 * Sets wether to delete file
	 * @param bool $delete
	 * @return Reader
	 */
	public function deleteFileOnFinish($delete = true) {
		$this->deleteFileOnFinish = $delete;
		return $this;
	}

	/**
	 *
	 * @return int
	 */
	public function key() {
		return $this->currentRow;
	}

	/**
	 *
	 * @return int
	 */
	public function next() {
		$this->currentRow++;
	}

	/**
	 *
	 * @return int
	 */
	public function rewind() {
		$this->currentRow = $this->firstDataRow;
	}

	/**
	 *
	 * @return bool
	 */
	public function valid() {
		$data = $this->current();
		$filteredData = array_filter($data);
		return count($filteredData) > 0;
	}

	/**
	 *
	 * @return mixed[]
	 */
	public function current() {
		$values = [];
		foreach ($this->loweredFieldNameToExcelColumnMap as $name => $index) {
			$value = $this->activeSheet->getCell($index . $this->currentRow)->getCalculatedValue();
			$values[$name] = trim($value);
		}

		$filteredValues = array_filter($values);
		if (count($filteredValues) < 1) {
			return $filteredValues;
		}

		return $this->formatValues($values);
	}

	/**
	 *
	 * @throws PhpExcelException
	 */
	private function mapHeaders() {
		$keys = array_keys($this->loweredFields);
		$columns = array_fill_keys($keys, null);
		$this->loweredFieldNameToExcelColumnMap = [];

		$lastRow = $this->activeSheet->getHighestRow();
		for ($i = $this->headerRow; $i <= $lastRow; $i++) {
			foreach ($this->excelColumnsRange as $columnIndex) {
				$value = $this->activeSheet->getCell($columnIndex . $i)->getCalculatedValue();
				$text = $this->lowerHeaderCellText($value);
				if (array_key_exists($text, $columns)) {
					$columns[$text] = $columnIndex;
				}
			}

			$this->loweredFieldNameToExcelColumnMap = array_filter($columns);
			if (count($this->loweredFieldNameToExcelColumnMap) > 0) {
				$this->firstDataRow = $i + 1;
				break;
			}
		}

		$missingColumns = array_diff_key($this->loweredToOriginalKeysMap, $this->loweredFieldNameToExcelColumnMap);
		if (count($missingColumns) > 0) {
			throw new PhpExcelException('Missing columns: ' . implode(', ', $missingColumns));
		}
	}

	/**
	 *
	 * @param array $values
	 * @return mixed[]
	 * @throws PhpExcelException
	 */
	private function formatValues(array $values) {
		$formattedValues = [];
		foreach ($values as $name => $value) {
			$originalFieldName = $this->loweredToOriginalKeysMap[$name];
			$type = $this->loweredFields[$name];

			if ($type & ExcelFieldType::REQUIRED) {
				if (strlen($value) < 1) {
					throw new PhpExcelException("Required column '$originalFieldName' is empty");
				}
				$type = $type ^ ExcelFieldType::REQUIRED;
			}

			if (strlen($value) < 1) {
				$formattedValues[$originalFieldName] = null;
				continue;
			}

			switch ($type) {
				case ExcelFieldType::DATE:
					$value = $this->getValueAsDate($value, $originalFieldName);
					break;

				case ExcelFieldType::INT:
					$value = $this->getValueAsInt($value, $originalFieldName);
					break;

				case ExcelFieldType::FLOAT:
					$value = $this->getValueAsFloat($value, $originalFieldName);
					break;
			}

			$formattedValues[$originalFieldName] = is_object($value) ? $value : StringHelper::safeTrim($value);
		}

		return $formattedValues;
	}

	/**
	 *
	 * @param mixed $value
	 * @param string $originalFieldName
	 * @return \DateTime|null
	 * @throws PhpExcelException
	 */
	private function getValueAsDate($value, $originalFieldName) {
		if (strlen($value) > 0 && !is_numeric($value)) {
			throw new PhpExcelException(
				"Invalid date value '$value'. "
				. "Make sure the cell is formatted as date. "
				. "Field '$originalFieldName', "
				. "Row '" . $this->currentRow . "'"
			);
		}
		return strlen($value) > 0 ? \PHPExcel_Shared_Date::ExcelToPHPObject($value) : null;
	}

	/**
	 *
	 * @param mixed $value
	 * @param string $originalFieldName
	 * @return int
	 * @throws PhpExcelException
	 */
	private function getValueAsInt($value, $originalFieldName) {
		if (strlen($value) > 0 && !preg_match('#^-?[0-9]*$#', $value)) {
			throw new PhpExcelException(
				"Invalid integer value '$value'. "
				. "Field '$originalFieldName', "
				. "Row '" . $this->currentRow . "'"
			);
		}

		return (int) $value;
	}

	/**
	 *
	 * @param mixed $value
	 * @param string $originalFieldName
	 * @return float
	 * @throws PhpExcelException
	 */
	private function getValueAsFloat($value, $originalFieldName) {
		if (strlen($value) > 0 && !is_numeric($value)) {
			throw new PhpExcelException(
				"Invalid float value '$value'. "
				. "Field '$originalFieldName', "
				. "Row '" . $this->currentRow . "'"
			);
		}

		return (float) $value;
	}

	/**
	 *
	 * @param string $text
	 * @return string
	 */
	private function lowerHeaderCellText($text) {
		return str_replace(' ', '', trim(strtolower($text)));
	}

}
