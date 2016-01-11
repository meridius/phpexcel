<?php

namespace Meridius\PhpExcel;

use \PHPExcel as PhpOffice_PHPExcel;
use \PHPExcel_Worksheet as PhpOffice_PHPExcel_Worksheet;
use \PHPExcel_Exception as PhpOffice_PHPExcel_Exception;
use Meridius\PhpExcel\Formatter;

class Worksheet extends \Nette\Object {

	/** @var PhpOffice_PHPExcel_Worksheet */
	private $sheet;

	/** @var Formatter */
	private $formatter;

	public function __construct(PhpOffice_PHPExcel $excel) {
		$this->sheet = $excel->createSheet();
		$this->formatter = new Formatter($this->sheet);
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
	 * @return PhpOffice_PHPExcel_Worksheet
	 */
	public function setTitle($pValue = 'Worksheet', $updateFormulaCellReferences = true) {
		$this->sheet->setTitle($pValue, $updateFormulaCellReferences);
	}

	/**
	 * Fill worksheet from values in array
	 *
	 * @param array $source Source array
	 * @param mixed $nullValue Value in source array that stands for blank cell
	 * @param string $startCell Insert array starting from this cell address as the top left coordinate
	 * @param boolean $strictNullComparison Apply strict comparison when testing for null values in the array
	 * @throws PhpOffice_PHPExcel_Exception
	 * @return PhpOffice_PHPExcel_Worksheet
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
	 * @param string $dataRange In format A2:E30
	 */
	public function applyStandardSheetFormat($dataRange) {
		$this->formatter->applyStandardSheetFormat($dataRange);
	}

}
