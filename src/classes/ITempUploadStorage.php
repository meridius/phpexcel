<?php

namespace Meridius\PhpExcel;

interface ITempUploadStorage {

	/**
	 * Get absolute path for given file name
	 * @param string $fileName
	 * @return string full path
	 */
	public function getFullPath($fileName);

}
