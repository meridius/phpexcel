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

}
