<?php

namespace Meridius\PhpExcel;

use Nette\Object;

class Dimension extends Object {

	/** @var Coordinate */
	private $topLeft;

	/** @var Coordinate */
	private $bottomRight;

	/**
	 *
	 * @param string $dimensionString in format A1:C5
	 * @throws PhpExcelException
	 */
	public function __construct($dimensionString) {
		$dimensionStringUpper = $this->checkDimension($dimensionString);
		$dimension = $this->separateDimension($dimensionStringUpper);
		$this->topLeft = $dimension[0];
		$this->bottomRight = $dimension[1];
	}

	public function __toString() {
		return $this->topLeft . ':' . $this->bottomRight;
	}

	/**
	 * @return Coordinate
	 */
	public function getTopLeftCoordinate() {
		return $this->topLeft;
	}

	/**
	 * @return Coordinate
	 */
	public function getTopRightCoordinate() {
		return new Coordinate($this->getLastColName() . $this->getFirstRowNum());
	}

	/**
	 * @return Coordinate
	 */
	public function getBottomLeftCoordinate() {
		return new Coordinate($this->getFirstColName() . $this->getLastRowNum());
	}

	/**
	 * @return Coordinate
	 */
	public function getBottomRightCoordinate() {
		return $this->bottomRight;
	}

	/**
	 * Name of the leftmost column
	 * @return string in format A
	 */
	public function getFirstColName() {
		return $this->getTopLeftCoordinate()->getColName();
	}

	/**
	 * Number of the first row
	 * @return int
	 */
	public function getFirstRowNum() {
		return $this->getTopLeftCoordinate()->getRowNum();
	}

	/**
	 * Name of the rightmost column
	 * @return string in format A
	 */
	public function getLastColName() {
		return $this->getBottomRightCoordinate()->getColName();
	}

	/**
	 * Number of the last row
	 * @return int
	 */
	public function getLastRowNum() {
		return $this->getBottomRightCoordinate()->getRowNum();
	}

	/**
	 *
	 * @return Coordinate[] [top left coordinate, bottom right coordinate]
	 */
	private function separateDimension($dimensionString) {
		$coordinates = explode(':', $dimensionString);
		return [
			new Coordinate($coordinates[0]),
			new Coordinate($coordinates[1]),
		];
	}

	/**
	 *
	 * @param string $dimensionString
	 * @return string
	 * @throws PhpExcelException
	 */
	private function checkDimension($dimensionString) {
		if (!is_string($dimensionString)) {
			throw new PhpExcelException('Only string can be used as dimension.');
		}
		$dimensionStringUpper = strtoupper($dimensionString);
		if (!preg_match('/^[A-Z]+[1-9]+\d*:[A-Z]+[1-9]+\d*$/i', $dimensionStringUpper)) {
			throw new PhpExcelException("Invalid dimension format '$dimensionString'.");
		}
		return $dimensionStringUpper;
	}

}
