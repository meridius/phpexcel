<?php

namespace Meridius\PhpExcel;

class Writer extends \Nette\Object {

	/**
	 *
	 * @return \Meridius\PhpExcel\Workbook
	 */
	public static function createNew() {
		return new Workbook;
	}

	/**
	 *
	 * @param string $filePath
	 * @return \Meridius\PhpExcel\Workbook
	 */
	public static function load($filePath) {
		return new Workbook($filePath, false);
	}

}
