<?php

namespace MeridiusTests\PhpExcel\Reader;

use Meridius\PhpExcel\AbstractPhpExcelReader;
use Meridius\PhpExcel\ExcelFieldType;
use Meridius\PhpExcel\Reader;
use MeridiusTests\PhpExcel\ExcelEntity\TestFileExcelEntity;
use MeridiusTests\PhpExcel\TempUploadStorage;

class TestFileAddReader extends AbstractPhpExcelReader {

	public function __construct(TempUploadStorage $tempUploadStorage) {
		parent::__construct($tempUploadStorage, new TestFileExcelEntity);
	}

	/**
	 *
	 * @param string $fileName
	 * @return TestFileExcelEntity[]
	 */
	public function readAddedList($fileName) {
		$reader = new Reader($fileName);
		$reader
			->setSheetNameToRead('Added list')
			->setExcelColumnsRange('A', 'D')
			->setFieldsToRead([
				TestFileExcelEntity::COL_1 => ExcelFieldType::FLOAT,
				TestFileExcelEntity::COL_2 => ExcelFieldType::DATE,
				TestFileExcelEntity::COLUMN_3 => ExcelFieldType::STRING,
				TestFileExcelEntity::COLUMN_4 => ExcelFieldType::INT | ExcelFieldType::REQUIRED,
			])
			->open();

		return $this->toEntities($reader);
	}

	/**
	 *
	 * @param string $fileName
	 * @return TestFileExcelEntity[]
	 */
	public function readOldList($fileName) {
		$reader = new Reader($fileName);
		$reader
			->setSheetNameToRead('List1')
			->setExcelColumnsRange('A', 'D')
			->setFieldsToRead([
				TestFileExcelEntity::COL_1 => ExcelFieldType::FLOAT,
				TestFileExcelEntity::COL_2 => ExcelFieldType::DATE,
				TestFileExcelEntity::COLUMN_3 => ExcelFieldType::STRING,
				TestFileExcelEntity::COLUMN_4 => ExcelFieldType::INT | ExcelFieldType::REQUIRED,
			])
			->open();

		return $this->toEntities($reader);
	}

}
