<?php

namespace MeridiusTests\PhpExcel;

use Meridius\PhpExcel\ITempUploadStorage;

class TempUploadStorage extends \Nette\Object implements ITempUploadStorage {

	/** @var string */
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
