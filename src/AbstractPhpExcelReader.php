<?php

namespace Meridius\PhpExcel;

use Nette\Object;

abstract class AbstractPhpExcelReader extends Object {

	/** @var ITempUploadStorage */
	private $tempUploadStorage;

	/** @var AbstractExcelEntity */
	private $blankEntity;

	/**
	 *
	 * @param ITempUploadStorage $tempUploadStorage
	 * @param AbstractExcelEntity $blankEntity
	 */
	public function __construct(ITempUploadStorage $tempUploadStorage, AbstractExcelEntity $blankEntity) {
		$this->tempUploadStorage = $tempUploadStorage;
		$this->blankEntity = $blankEntity;
	}

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
