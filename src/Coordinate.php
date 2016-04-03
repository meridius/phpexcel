<?php

namespace Meridius\PhpExcel;

use Meridius\Helpers\ExcelHelper;
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
		$coordinateStringUpper = $this->checkCoordinate($coordinateString);
		$coordinate = $this->separateCoordinate($coordinateStringUpper);
		$this->col = $coordinate[0];
		$this->row = $coordinate[1];
	}

	/**
	 * Coordinate in format A1
	 * @return string
	 */
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
	 *
	 * @param string $colName
	 * @return Coordinate
	 */
	public function setColName($colName) {
		$this->col = $this->checkColName($colName);
		return $this;
	}

	/**
	 *
	 * @param int $rowNum
	 * @return Coordinate
	 * @throws PhpExcelException
	 */
	public function setRowNum($rowNum) {
		$this->row = $this->checkNumber($rowNum);
		return $this;
	}

	/**
	 * How many columns add/remove to coordinate
	 * @param int $numCols positive value will add, negative will remove
	 * @return Coordinate
	 * @throws PhpExcelException
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
	private function checkCoordinate($coordinateString) {
		$coordinateStringUpper = $this->checkString($coordinateString, 'coordinate');
		if (!preg_match('/^[A-Z]+[1-9]+\d*$/', $coordinateStringUpper)) {
			throw new PhpExcelException("Invalid coordinate format '$coordinateString'.");
		}
		return $coordinateStringUpper;
	}

	/**
	 *
	 * @param string $colName
	 * @return string col name in uppercase
	 * @throws PhpExcelException
	 */
	private function checkColName($colName) {
		$colNameUpper = $this->checkString($colName, 'column name');
		if (!preg_match('/^[A-Z]+$/', $colNameUpper)) {
			throw new PhpExcelException('Invalid column name given.');
		}
		return $colNameUpper;
	}

	/**
	 *
	 * @param string $param
	 * @param string $where
	 * @return string
	 * @throws PhpExcelException
	 */
	private function checkString($param, $where) {
		if (!is_string($param)) {
			throw new PhpExcelException("Only string can be used as $where.");
		}
		return strtoupper($param);
	}

	/**
	 *
	 * @param int $param
	 * @return int
	 * @throws PhpExcelException
	 */
	private function checkNumber($param) {
		if (ctype_digit($param) || is_int($param)) {
			$param = (int) $param;
			if ($param > 0) {
				return $param;
			}
		}
		throw new PhpExcelException('Invalid row number given.');
	}

}
