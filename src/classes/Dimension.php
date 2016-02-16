<?php

namespace Meridius\PhpExcel;

use Nette\Object;

class Dimension extends Object {

	/** @var string */
	private static $dimensionString;

	/**
	 * 
	 * @param string $dimensionString in format A1:C5
	 * @throws PhpExcelException
	 */
	public function __construct($dimensionString) {
		if (!is_string($dimensionString)) {
			throw new PhpExcelException('Only string can be used as dimension.');
		}
		if (!preg_match('/^[a-z]+\d+:[a-z]+\d+$/i', $dimensionString)) {
			throw new PhpExcelException("Invalid dimension format '$dimensionString'.");
		}
		self::$dimensionString = strtoupper($dimensionString);
	}
	
	/**
	 * @return string in format A1
	 */
	public static function getTopLeftCoordinate() {
		return self::separateDimension()[0];
	}

	/**
	 * @return string in format A1
	 */
	public static function getTopRightCoordinate() {
		return self::getLastColName() . self::getFirstRowNum();
	}
	
	/**
	 * @return string in format A1
	 */
	public static function getBottomRightCoordinate() {
		return self::separateDimension()[1];
	}

	/**
	 * @return string in format A1
	 */
	public static function getBottomLeftCoordinate() {
		return self::getFirstColName() . self::getLastRowNum();
	}
	
	/**
	 * Name of the leftmost column
	 * @return string in format A
	 */
	public static function getFirstColName() {
		$coordinate = self::getTopLeftCoordinate();
		return self::separateCoordinate($coordinate)[0];
	}
	
	/**
	 * Number of the first row
	 * @return int
	 */
	public static function getFirstRowNum() {
		$coordinate = self::getTopLeftCoordinate();
		return (int) self::separateCoordinate($coordinate)[1];
	}
	
	/**
	 * Name of the rightmost column
	 * @return string in format A
	 */
	public static function getLastColName() {
		$coordinate = self::getBottomRightCoordinate();
		return self::separateCoordinate($coordinate)[0];
	}
	
	/**
	 * Number of the last row
	 * @return int
	 */
	public static function getLastRowNum() {
		$coordinate = self::getBottomRightCoordinate();
		return (int) self::separateCoordinate($coordinate)[1];
	}
	
	/**
	 * 
	 * @return array [top left coordinate, bottom right coordinate]
	 */
	private static function separateDimension() {
		return explode(':', self::$dimensionString);
	}
	
	/**
	 * 
	 * @param string $coordinate
	 * @return array [column name, row number]
	 */
	private static function separateCoordinate($coordinate) {
		$matches = [];
		preg_match('/^([a-z]+)(\d+)$/i', $coordinate, $matches);
		array_shift($matches);
		return $matches;
	}

}