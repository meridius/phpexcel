<?php

namespace Meridius\PhpExcel;

use \PHPExcel_Worksheet as PhpOffice_PHPExcel_Worksheet;
use \PHPExcel_Exception as PhpOffice_PHPExcel_Exception;
use Meridius\PhpExcel\Formatter;

class Worksheet extends \Nette\Object {

	/** @var PhpOffice_PHPExcel_Worksheet */
	private $sheet;

	/** @var Formatter */
	private $formatter;

	public function __construct(PhpOffice_PHPExcel_Worksheet $sheet) {
		$this->sheet = $sheet;
		$this->formatter = new Formatter($this->sheet);
	}

	/**
	 *
	 * @return PhpOffice_PHPExcel_Worksheet
	 */
	public function getPhpOfficeWorksheetObject() {
		return $this->sheet;
	}

	/**
	 * Will return dimension of data in worksheet
	 * @return Dimension
	 */
	public function getDataDimension() {
		$dimensionsString = $this->sheet->calculateWorksheetDataDimension();
		return new Dimension($dimensionsString);
	}

	/**
	 * Set title
	 *
	 * @param string $pValue String containing the dimension of this worksheet
	 * @param string $updateFormulaCellReferences boolean Flag indicating whether cell references in formulae should
	 * be updated to reflect the new sheet name.
	 * This should be left as the default true, unless you are
	 * certain that no formula cells on any worksheet contain
	 * references to this worksheet
	 * @return Worksheet
	 */
	public function setTitle($pValue = 'Worksheet', $updateFormulaCellReferences = true) {
		$this->sheet->setTitle($pValue, $updateFormulaCellReferences);
		return $this;
	}

	/**
	 * Set a cell value
	 *
	 * @param string $pCoordinate Coordinate of the cell
	 * @param mixed $pValue Value of the cell
	 * @return Worksheet
	 */
	public function setCellValue($pCoordinate = 'A1', $pValue = null) {
		$this->sheet->setCellValue($pCoordinate, $pValue);
		return $this;
	}

	/**
	 * Fill worksheet from values in array
	 *
	 * @param array $source Source array
	 * @param mixed $nullValue Value in source array that stands for blank cell
	 * @param string $startCell Insert array starting from this cell address as the top left coordinate
	 * @param boolean $strictNullComparison Apply strict comparison when testing for null values in the array
	 * @throws PhpOffice_PHPExcel_Exception
	 * @return Worksheet
	 */
	public function fromArray($source = null, $nullValue = null, $startCell = 'A1', $strictNullComparison = false) {
		foreach ($source as &$row) {
			if (is_array($row)) {
				foreach ($row as $key => $value) {
					if ($value instanceof \Nette\Utils\DateTime) {
						$row[$key] = \PHPExcel_Shared_Date::PHPToExcel($value);
					}
				}
			}
		}
		$this->sheet->fromArray($source, $nullValue, $startCell, $strictNullComparison);
		return $this;
	}

	/**
	 *
	 * @return Formatter
	 */
	public function getFormatter() {
		return $this->formatter;
	}

	/**
	 * Apply standard formating for sheet
	 * @param string|null $dataDimension In format A2:E30 or automatically by data dimension
	 * @return Worksheet
	 */
	public function applyStandardSheetFormat($dataDimension = null) {
		if (!$dataDimension) {
			$dataDimension = (string) $this->getDataDimension();
		}
		$this->formatter->applyStandardSheetFormat($dataDimension);
		return $this;
	}

}
