<?php

namespace services\storages;

class TempExcelStorage extends \Nette\Object {

	const PHPEXCEL_2007 = 'Excel2007';
	const PHPEXCEL_2003 = 'Excel5';

	/** @var string */
	private $directory;

	/**
	 *
	 * @param string $tempDir
	 */
	public function __construct($tempDir) {
		$this->directory = $tempDir . DIRECTORY_SEPARATOR . 'excel';
		if (!is_dir($this->directory)) {
			mkdir($this->directory);
		}
	}

	/**
	 *
	 * @param \COMExcel\Workbook $excel
	 * @param bool $returnFileNameOnly
	 * @return string
	 */
	public function saveComExcel(\COMExcel\Workbook $excel, $returnFileNameOnly = false) {
		$tempName = md5(uniqid()) . '.xlsx';
		$path = $this->directory . DIRECTORY_SEPARATOR . $tempName;
		$excel->saveWorkbook($path);
		$excel->close();
		unset($excel);
		if ($returnFileNameOnly) {
			return $tempName;
		}
		return $path;
	}

	/**
	 * Saves excel to temp storage and returns path to saved file
	 * @param \PHPExcel $excel
	 * @param boolean $returnFileNameOnly
	 * @param string $type xls|xlsx
	 * @return string path to saved file
	 * @throws \exceptions\PrivateException
	 */
	public function savePhpExcel(\PHPExcel $excel, $returnFileNameOnly = false, $type = 'xlsx') {
		switch ($type) {
			case 'xls':
				$writerType = self::PHPEXCEL_2003;
				break;
			case 'xlsx':
				$writerType = self::PHPEXCEL_2007;
				break;
			default:
				throw new \exceptions\PrivateException("Invalid excel type '$type'.");
		}
		$writer = \PHPExcel_IOFactory::createWriter($excel, $writerType);
		$tempName = md5(uniqid()) . '.' . $type;
		$path = $this->directory . DIRECTORY_SEPARATOR . $tempName;
		$writer->save($path);
		return $returnFileNameOnly ? $tempName : $path;
	}

	/**
	 * Saves excel to temp storage and returns path to saved file
	 * @param \Meridius\PhpExcel\Workbook $excel
	 * @param boolean $returnFileNameOnly
	 * @param string $type xls|xlsx
	 * @return string path to saved file
	 * @throws \exceptions\PrivateException
	 */
	public function savePhpExcelWorkbook(\Meridius\PhpExcel\Workbook $excel, $returnFileNameOnly = false, $type = 'xlsx') {
		$tempName = md5(uniqid()) . '.' . $type;
		$path = $this->directory . DIRECTORY_SEPARATOR . $tempName;
		$excel->save($path, $type);
		return $returnFileNameOnly ? $tempName : $path;
	}

	/**
	 * Saves uploaded excel to temporary storage. Usable for <b>PHP Excel</b>.<br/>
	 * the Excel COM objects couldn't access the file.</i>
	 * @param \Nette\Http\FileUpload $file
	 * @return string new file path
	 */
	public function saveUploadedByMoving(\Nette\Http\FileUpload $file) {
		$extension = $this->getExtension($file);
		$tempName = $this->directory . DIRECTORY_SEPARATOR . md5(uniqid()) . '.' . $extension;

		$file->move($tempName);

		return $tempName;
	}

	/**
	 * Saves uploaded excel to temporary storage. Usable for <b>COM Excel</b>.<br/>
	 * <i>Fix so Excel COM Object can open uploaded file <br/>
	 * but it breaks opening Excel files in PHPExcel.<br/>
	 * Using move() will copy user rights too - we don't want that in this case!!!</i>
	 * @param \Nette\Http\FileUpload $file
	 * @return string new file path
	 */
	public function saveUploadedByCopying(\Nette\Http\FileUpload $file) {
		$extension = $this->getExtension($file);
		$tempName = $this->directory . DIRECTORY_SEPARATOR . md5(uniqid()) . '.' . $extension;

		copy($file, $tempName);
		unlink($file);

		return $tempName;
	}

	/**
	 *
	 * @param \Nette\Http\FileUpload $file
	 * @return string file extension
	 */
	public function getExtension(\Nette\Http\FileUpload $file) {
		$matches = [];
		preg_match('#\.([^\.]+)$#', $file->getName(), $matches);
		return isset($matches[1]) ? $matches[1] : 'xls';
	}

	/**
	 * Returns full path to a file based on a name
	 * @param string $fileName
	 * @throws \Nette\IOException
	 */
	public function getFileByName($fileName) {
		$fullPath = $this->directory . DIRECTORY_SEPARATOR . $fileName;
		if (!is_file($fullPath)) {
			throw new \Nette\IOException('Requested file was not found');
		}

		return $fullPath;
	}

	/**
	 * Deleters temporary files that are older than specified date
	 * @param \DateTime $date
	 */
	public function deleteFilesOlderThan(\DateTime $date) {
		$directory = new \DirectoryIterator($this->directory);
		foreach ($directory as $file) {
			if (!$file->isFile()) {
				continue;
			}
			$time = $file->getMTime();

			if ($date->getTimestamp() > $time) {
				unlink($file->getPathname());
			}
		}
	}

}
