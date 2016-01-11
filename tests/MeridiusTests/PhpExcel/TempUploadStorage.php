<?php

namespace services\storages;

use Meridius\PhpExcel\ITempUploadStorage;

class TempUploadStorage extends \Nette\Object implements ITempUploadStorage {

	/**
	 *
	 * @var String
	 */
	private $realPath;

	/**
	 *
	 * @param string $uploadTempDir
	 */
	public function __construct($uploadTempDir) {
		$this->realPath = $uploadTempDir;
	}

	/**
	 *
	 * @param string $file
	 * @return string
	 */
	public function getFullPath($file) {
		return $this->realPath . DIRECTORY_SEPARATOR . $file;
	}

}
