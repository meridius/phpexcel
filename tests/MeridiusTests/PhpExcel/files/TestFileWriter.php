<?php

namespace MeridiusTests\PhpExcel\Reader;

use Meridius\PhpExcel\Workbook;
use Meridius\PhpExcel\Writer;
use MeridiusTests\PhpExcel\ExcelEntity\TestFileExcelEntity;
use MeridiusTests\PhpExcel\TempExcelStorage;

class TestFileWriter {

	/** @var Workbook */
	private $excel;

	/** @var TempExcelStorage */
	private $tempExcelStorage;

	public function __construct(TempExcelStorage $tempExcelStorage) {
		$this->tempExcelStorage = $tempExcelStorage;
	}

	/**
	 *
	 * @param TestFileExcelEntity[] $data
	 * @return string file path
	 */
	public function writeFile(array $data) {
		$this->excel = Writer::createNew();
		$this->writeFirstSheet($data);
		return $this->tempExcelStorage->savePhpExcelWorkbook($this->excel);
	}

	/**
	 *
	 * @param TestFileExcelEntity[] $data
	 */
	private function writeFirstSheet(array $data) {
		$header = [
			TestFileExcelEntity::COL_1,
			TestFileExcelEntity::COL_2,
			TestFileExcelEntity::COLUMN_3,
			TestFileExcelEntity::COLUMN_4,
		];
		$dataArray = [];
		foreach ($data as $item) {
			$dataArray[] = [
				$item->col1,
				$item->col2,
				$item->column3,
				$item->column4,
			];
		}
		$this->excel->addBasicSheet($header, $dataArray, 'List1');
	}

}
