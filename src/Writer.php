<?php

namespace Meridius\PhpExcel;

class Writer extends \Nette\Object {

	public static function createNew() {
		return new Workbook;
	}

}
