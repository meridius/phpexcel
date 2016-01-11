<?php

namespace services\readers;

use Meridius\PhpExcel\AbstractPhpExcelReader;
use Meridius\PhpExcel\Reader;
use Meridius\PhpExcel\ExcelFieldType;
use \ExcelEntities\AllSerialsReaderExcelEntity;

class ReaderTest2 extends AbstractPhpExcelReader {

	public function __construct(\services\storages\TempUploadStorage $tempUploadStorage) {
		parent::__construct($tempUploadStorage, new AllSerialsReaderExcelEntity);
	}

	/**
	 *
	 * @param string $fileName
	 * @return AllSerialsReaderExcelEntity[]
	 */
	public function readFile($fileName) {
		$fullPath = $this->getFullPath($fileName);
		$reader = new Reader($fullPath);
		$reader
			->setSheetNameToRead('List1')
			->setExcelColumnsRange('A', 'D')
			->setFieldsToRead([
				AllSerialsReaderExcelEntity::COL_1 => ExcelFieldType::FLOAT,
				AllSerialsReaderExcelEntity::COL_2 => ExcelFieldType::DATE,
				AllSerialsReaderExcelEntity::COLUMN_3 => ExcelFieldType::STRING,
				AllSerialsReaderExcelEntity::COLUMN_4 => ExcelFieldType::INT | ExcelFieldType::REQUIRED,
			])
//			->deleteFileOnFinish()
			->open();

		return $this->toEntities($reader);
	}

}
