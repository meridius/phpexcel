<?php

namespace Meridius\PhpExcel;

use \Meridius\PhpExcel\PhpExcelException;

abstract class AbstractExcelEntity extends \Nette\Object {

	/**
	 *
	 * @param array $row
	 */
	public function __construct($row = null) {
		if (!is_null($row)) {
			$this->loadFromExcelRow($row);
		}
	}

	/**
	 * @return array
	 */
	abstract protected function getMappingArray();

	/**
	 *
	 * @param array $row
	 * @return AbstractExcelEntity
	 */
	public function loadFromExcelRow($row) {
		$mappingArray = $this->getMappingArray();
		foreach ($mappingArray as $column => &$varRef) {
			if (array_key_exists($column, $row)) {
				$varRef = $row[$column];
			}
		}
		return $this;
	}

	/**
	 * Checks if property is not null.
	 * @param string $constName
	 * @return bool
	 * @throws PhpExcelException
	 */
	public function isPropDefined($constName) {
		return !is_null($this->getByConst($constName));
	}

	/**
	 * Get value from mapping array by constant
	 * @param string $constName
	 * @return mixes
	 */
	public function getByConst($constName) {
		$mappingArray = $this->checkConstValidity($constName);
		return $mappingArray[$constName];
	}

	/**
	 *
	 * @param string $constName
	 * @return array mapping array
	 * @throws PhpExcelException
	 */
	private function checkConstValidity($constName) {
		$mappingArray = $this->getMappingArray();
		if (!array_key_exists($constName, $mappingArray)) {
			$class = $this->getReflection()->getName();
			throw new PhpExcelException("Constant '$constName' is not defined in '$class'.");
		}
		return $mappingArray;
	}

}
