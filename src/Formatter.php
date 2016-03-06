<?php

namespace Meridius\PhpExcel;

use Meridius\PhpExcel\PhpExcelException;
use Nette\Object;
use PHPExcel_Cell as PhpOffice_PHPExcel_Cell;
use PHPExcel_Cell_DataValidation as PhpOffice_PHPExcel_Cell_DataValidation;
use PHPExcel_Style_Alignment as PhpOffice_PHPExcel_Style_Alignment;
use PHPExcel_Style_Fill as PhpOffice_PHPExcel_Style_Fill;
use PHPExcel_Style_Protection as PhpOffice_PHPExcel_Style_Protection;
use PHPExcel_Worksheet as PhpOffice_PHPExcel_Worksheet;

class Formatter extends Object {

	const DATE_FORMAT_EXCEL = 'd-mmm-yyyy';
	const TIME_FORMAT_EXCEL = 'h:mm';
	const DATETIME_FORMAT_EXCEL = 'd-mmm-yyyy (h:mm)';

	/** @var PhpOffice_PHPExcel_Worksheet */
	private $sheet;

	public function __construct(PhpOffice_PHPExcel_Worksheet $sheet) {
		$this->sheet = $sheet;
	}

	/**
	 * Apply standard formating for sheet
	 * @param string $dataRange In format A2:E30
	 * @return Formatter
	 */
	public function applyStandardSheetFormat($dataRange) {
		$this->autoFilter($dataRange);
		$this->autosizeColumns();
		$this->freezePane();
		$this->setSelectedCell();
		return $this;
	}

	/**
	 * Lock whole sheet. All cells are protected.
	 * @param string $password default null
	 * @return Formatter
	 */
	public function lockSheet($password = null) {
		$this->sheet->getProtection()->setSheet(true);
		if (!is_null($password)) {
			$this->sheet->getProtection()->setPassword($password);
		}
		return $this;
	}

	/**
	 * Unlock range in locked sheet.
	 * @param string $range
	 * @return Formatter
	 */
	public function unlockRange($range) {
		$this->sheet->getStyle($range)->getProtection()->setLocked(
			PhpOffice_PHPExcel_Style_Protection::PROTECTION_UNPROTECTED
		);
		return $this;
	}

	/**
	 * Set date format for column
	 * @param string $columnName
	 * @return Formatter
	 */
	public function formatColumnAsDate($columnName) {
		$this->formatColumn($columnName, self::DATE_FORMAT_EXCEL);
		return $this;
	}

	/**
	 * Set datetime format for column
	 * @param string $columnName
	 * @return Formatter
	 */
	public function formatColumnAsDatetime($columnName) {
		$this->formatColumn($columnName, self::DATETIME_FORMAT_EXCEL);
		return $this;
	}

	/**
	 * Set format for column
	 * @param string $columnName
	 * @return Formatter
	 */
	private function formatColumn($columnName, $formatString) {
		$this->sheet
			->getStyle($columnName . '1:' . $columnName . $this->sheet->getHighestDataRow())
			->getNumberFormat()
			->setFormatCode($formatString);
		return $this;
	}

	/**
	 * Set excel validation for cell range
	 * @param string $range In format A2:E30
	 * @param string $sourceDataFormula In format Sites!$A$2:$A$4233
	 * @param string $errorMessage Optional
	 * @return Formatter
	 * @throws PhpExcelException
	 */
	public function setDataValidationRange($range, $sourceDataFormula, $errorMessage = 'Invalid data') {
		$regCoor = '([a-zA-Z]+)(\d+)';
		if (!preg_match("/$regCoor:$regCoor/", $range)) {
			throw new PhpExcelException('Range is not in valid format');
		}
		list($fromCoordinates, $toCoordinates) = explode(':', $range);
		$matches = [];
		preg_match("/$regCoor/", $fromCoordinates, $matches);
		list(, $fromColName, $fromRowNum) = $matches; // skip $matches[0]
		preg_match("/$regCoor/", $toCoordinates, $matches);
		list(, $toColName, $toRowNum) = $matches; // skip $matches[0]

		$fromColNum = PhpOffice_PHPExcel_Cell::columnIndexFromString($fromColName); // A = 1
		$toColNum = PhpOffice_PHPExcel_Cell::columnIndexFromString($toColName);

		if ($fromColNum > $toColNum || $fromRowNum > $toRowNum) {
			throw new PhpExcelException("Range '$range' is not in valid format");
		}

		for ($colNum = $fromColNum; $colNum <= $toColNum; $colNum++) {
			$colName = PhpOffice_PHPExcel_Cell::stringFromColumnIndex($colNum - 1); // A = 0
			for ($rowNum = $fromRowNum; $rowNum <= $toRowNum; $rowNum++) {
				$objValidation = $this->sheet->getDataValidation($colName . $rowNum);
				$objValidation->setType(PhpOffice_PHPExcel_Cell_DataValidation::TYPE_LIST);
				$objValidation->setErrorStyle(PhpOffice_PHPExcel_Cell_DataValidation::STYLE_INFORMATION);
				$objValidation->setAllowBlank(false);
				$objValidation->setShowInputMessage(true);
				$objValidation->setShowErrorMessage(true);
				$objValidation->setShowDropDown(true);
				$objValidation->setErrorTitle('Input error');
				$objValidation->setError($errorMessage);
				$objValidation->setFormula1($sourceDataFormula);
			}
		}
		return $this;
	}

	/**
	 *
	 * @param string $startColName
	 * @param string $endColName
	 * @return Formatter
	 */
	public function autosizeColumns($startColName = null, $endColName = null) {
		if (is_null($startColName) || is_null($endColName)) {
			list($startCoordinates, $endCoordinates) = explode(':', $this->sheet->calculateWorksheetDataDimension());
			$startColName = is_null($startColName) ? preg_replace('/\d*/', '', $startCoordinates) : $startColName;
			$endColName = is_null($endColName) ? preg_replace('/\d*/', '', $endCoordinates) : $endColName;
		}
		$startNum = PhpOffice_PHPExcel_Cell::columnIndexFromString($startColName); // A = 1
		$endNum = PhpOffice_PHPExcel_Cell::columnIndexFromString($endColName);

		for ($i = $startNum - 1; $i < $endNum; $i++) {
			$this->sheet->getColumnDimensionByColumn($i)->setAutoSize(true); // A = 0
		}
		$this->sheet->calculateColumnWidths();
		return $this;
	}

	/**
	 *
	 * @param string $range In format A2:E30
	 * @return Formatter
	 */
	public function autoFilter($range) {
		$this->sheet->setAutoFilter($range);
		return $this;
	}

	/**
	 *
	 * @param string $address In format A1
	 * @return Formatter
	 */
	public function freezePane($address = 'A2') {
		$this->sheet->freezePane($address);
		return $this;
	}

	/**
	 *
	 * @param string $address In format A1
	 * @return Formatter
	 */
	public function setSelectedCell($address = 'A1') {
		$this->sheet->setSelectedCell($address);
		return $this;
	}

	/**
	 *
	 * @param sring $range In format A2:E30
	 * @return Formatter
	 */
	public function setBackgroundColor($range, $rgb) {
		$this->sheet->getStyle($range)
			->applyFromArray(
				[
					'fill' => [
						'type' => PhpOffice_PHPExcel_Style_Fill::FILL_SOLID,
						'color' => ['rgb' => $rgb],
					],
				]
			);
		return $this;
	}

	/**
	 *
	 * @param integer $height
	 * @return Formatter
	 */
	public function setRowsHeight($height = -1) {
		foreach ($this->sheet->getRowDimensions() as $rd) {
			$rd->setRowHeight($height);
		}
		return $this;
	}

	/**
	 *
	 * @param type $range
	 * @return Formatter
	 */
	public function mergeAndCenter($range) {
		$this->sheet->mergeCells($range);
		$this->sheet->getStyle($range)
			->getAlignment()
			->setHorizontal(PhpOffice_PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		return $this;
	}

}
