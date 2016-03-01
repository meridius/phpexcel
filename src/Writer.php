<?php

namespace Meridius\PhpExcel;

use Nette\Object;

class Writer extends Object {

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
