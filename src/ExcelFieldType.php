<?php

namespace Meridius\PhpExcel;

use Nette\Object;

class ExcelFieldType extends Object {

	const REQUIRED = 1;
	const STRING = 2;
	const FLOAT = 4;
	const DATE = 8;
	const INT = 16;

}
