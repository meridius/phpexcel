<?php

namespace Meridius\PhpExcel;

use Meridius\Helpers\ExcelHelper;
use Meridius\PhpExcel\PhpExcelException;
use Nette\Object;

class Coordinate extends Object {

	/** @var string */
	private $col;

	/** @var int */
	private $row;

	/**
	 *
	 * @param string $coordinateString in format A1
	 * @throws PhpExcelException
	 */
	public function __construct($coordinateString) {
		$coordinateStringUpper = $this->checkCoordinateValidity($coordinateString);
		$coordinate = $this->separateCoordinate($coordinateStringUpper);
		$this->col = $coordinate[0];
		$this->row = $coordinate[1];
	}

	public function __toString() {
		return $this->col . $this->row;
	}

	/**
	 *
	 * @return string
	 */
	public function getColName() {
		return $this->col;
	}

	/**
	 *
	 * @return int
	 */
	public function getRowNum() {
		return (int) $this->row;
	}

	/**
	 * How many columns add/remove to coordinate
	 * @param int $numCols positive value will add, negative will remove
	 * @return Coordinate
	 */
	public function shiftColBy($numCols = 0) {
		if (!is_int($numCols)) {
			throw new PhpExcelException('Only integer values are allowed.');
		}
		$colNum = ExcelHelper::getExcelColumnNumber($this->col) + $numCols;
		if ($colNum < 1) {
			throw new PhpExcelException('Number of rows would get below 1.');
		}
		$this->col = ExcelHelper::getExcelColumnName($colNum);
		return $this;
	}

	/**
	 * How many rows add/remove to coordinate
	 * @param int $numRows positive value will add, negative will remove
	 * @return Coordinate
	 * @throws PhpExcelException
	 */
	public function shiftRowBy($numRows = 0) {
		if (!is_int($numRows)) {
			throw new PhpExcelException('Only integer values are allowed.');
		}
		$rowNum = $this->row + $numRows;
		if ($rowNum < 1) {
			throw new PhpExcelException('Number of rows would get below 1.');
		}
		$this->row = $rowNum;
		return $this;
	}

	/**
	 *
	 * @param string $coordinate
	 * @return array [column name, row number]
	 */
	private function separateCoordinate($coordinate) {
		$matches = [];
		preg_match('/^([A-Z]+)(\d+)$/', $coordinate, $matches);
		array_shift($matches);
		return $matches;
	}

	/**
	 *
	 * @param string $coordinateString
	 * @return string
	 * @throws PhpExcelException
	 */
	private function checkCoordinateValidity($coordinateString) {
		if (!is_string($coordinateString)) {
			throw new PhpExcelException('Only string can be used as coordinate.');
		}
		$coordinateStringUpper = strtoupper($coordinateString);
		if (!preg_match('/^[A-Z]+\d+$/', $coordinateStringUpper)) {
			throw new PhpExcelException("Invalid coordinate format '$coordinateString'.");
		}
		return $coordinateStringUpper;
	}

}
