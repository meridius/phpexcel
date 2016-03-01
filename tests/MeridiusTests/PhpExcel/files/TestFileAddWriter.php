<?php

namespace MeridiusTests\PhpExcel\Reader;

use Meridius\PhpExcel\Workbook;
use Meridius\PhpExcel\Writer;
use MeridiusTests\PhpExcel\ExcelEntity\TestFileExcelEntity;

class TestFileAddWriter {

	/** @var Workbook */
	private $excel;

	/**
	 *
	 * @param string $filePath
	 * @param TestFileExcelEntity[] $data
	 * @return Workbook
	 */
	public function writeToFile($filePath, array $data) {
		$this->excel = Writer::load($filePath);
		$this->addSecondSheet($data);
		return $this->excel;
	}

	/**
	 *
	 * @param TestFileExcelEntity[] $data
	 */
	private function addSecondSheet(array $data) {
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
				$item->col2 ? \PHPExcel_Shared_Date::PHPToExcel($item->col2) : null,
				$item->column3,
				$item->column4,
			];
		}
		$this->excel->addBasicSheet($header, $dataArray, 'Added list');
	}

}
