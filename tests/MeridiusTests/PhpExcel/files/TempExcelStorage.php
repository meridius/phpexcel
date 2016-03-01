<?php

namespace MeridiusTests\PhpExcel;

use Meridius\PhpExcel\Workbook;
use Nette\Object;

class TempExcelStorage extends Object {

	/** @var string */
	private $directory;

	/**
	 *
	 * @param string $tempDir
	 */
	public function __construct($tempDir) {
		$this->directory = "$tempDir/excel";
		if (!is_dir($this->directory)) {
			mkdir($this->directory);
		}
	}

	/**
	 * Saves excel to temp storage and returns path to saved file
	 * @param Workbook $excel
	 * @param boolean $returnFileNameOnly
	 * @param string $type
	 * @return string path to saved file
	 */
	public function savePhpExcelWorkbook(Workbook $excel, $returnFileNameOnly = false, $type = 'xlsx') {
		$tempName = md5(uniqid()) . '.' . $type;
		$path = $this->directory . '/' . $tempName;
		$excel->save($path, $type);
		return $returnFileNameOnly ? $tempName : $path;
	}

}
