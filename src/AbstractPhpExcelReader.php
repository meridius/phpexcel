<?php

namespace Meridius\PhpExcel;

use \Meridius\PhpExcel\Reader;
use \Meridius\PhpExcel\AbstractExcelEntity;
use \Meridius\PhpExcel\ITempUploadStorage;
use \Meridius\PhpExcel\PhpExcelException;

abstract class AbstractPhpExcelReader extends \Nette\Object {

	/** @var ITempUploadStorage */
	private $tempUploadStorage;

	/** @var AbstractExcelEntity */
	private $blankEntity;

	public function __construct(ITempUploadStorage $tempUploadStorage, AbstractExcelEntity $blankEntity) {
		$this->tempUploadStorage = $tempUploadStorage;
		$this->blankEntity = $blankEntity;
	}

	/**
	 * This method should create new instance of \Meridius\PhpExcel\Reader,<br/>
	 * set sheet name, columns, etc. and pass the reader to toEntities() method.
	 * It is now commented-out because you don't have to implement it in case separate methods for each sheet.
	 * @param string $file file path
	 * @return AbstractExcelEntity[]
	 */
// public function readFile($file);

	/**
	 *
	 * @param string $fileName
	 * @return string
	 */
	protected function getFullPath($fileName) {
		return $this->tempUploadStorage->getFullPath($fileName);
	}

	/**
	 *
	 * @param array $row
	 * @return AbstractExcelEntity
	 * @throws PhpExcelException
	 */
	protected function toEntity(array $row) {
		$entity = clone $this->blankEntity;
		return $entity->loadFromExcelRow($row);
	}

	/**
	 *
	 * @param Reader $rows
	 * @return AbstractExcelEntity[]
	 */
	protected function toEntities(Reader $rows) {
		$entities = [];
		foreach ($rows as $key => $row) {
			$entities[$key] = $this->toEntity($row);
		}
		return $entities;
	}

}
