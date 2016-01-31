<?php

namespace MeridiusTests\PhpExcel\ExcelEntity;

use Meridius\PhpExcel\AbstractExcelEntity;

class TestFileExcelEntity extends AbstractExcelEntity {

	const COL_1 = 'Col 1';
	const COL_2 = 'Col 2';
	const COLUMN_3 = 'Colum 3';
	const COLUMN_4 = 'Colum 4';

	public $col1;
	public $col2;
	public $column3;
	public $column4;

	/**
	 * @return array
	 */
	protected function getMappingArray() {
		return [
			self::COL_1 => &$this->col1,
			self::COL_2 => &$this->col2,
			self::COLUMN_3 => &$this->column3,
			self::COLUMN_4 => &$this->column4,
		];
	}

}
