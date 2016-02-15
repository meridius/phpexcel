<?php

namespace Meridius\PhpExcel;

use Meridius\Helpers\ExcelHelper;
use Meridius\PhpExcel\PhpExcelException;
use PHPExcel as PhpOffice_PHPExcel;
use PHPExcel_IOFactory as PhpOffice_PHPExcel_IOFactory;
use PHPExcel_Reader_Exception as PhpOffice_PHPExcel_Reader_Exception;
use PHPExcel_Writer_Exception as PhpOffice_PHPExcel_Writer_Exception;
use PHPExcel_Writer_IWriter as PhpOffice_PHPExcel_Writer_IWriter;

class Workbook extends \Nette\Object {

	const PHPEXCEL_2007 = 'Excel2007';
	const PHPEXCEL_2003 = 'Excel5';

	/** @var PhpOffice_PHPExcel */
	private $excel;

	/** @var string */
	private $dateFormat;

	/**
	 * Excel with one sheet
	 * @param string|null $filePath
	 * @param boolean $withoutSheets
	 */
	public function __construct($filePath = null, $withoutSheets = true) {
		try {
			$this->excel = $filePath
				? PhpOffice_PHPExcel_IOFactory::load($filePath)
				: new PhpOffice_PHPExcel;
		} catch (PhpOffice_PHPExcel_Reader_Exception $ex) {
			throw new PhpExcelException("Unable to create excel object.", 0, $e);
		}
		if (!$filePath && $withoutSheets) {
			$this->excel->removeSheetByIndex();
		}
	}

	/**
	 *
	 * @param array $header
	 * @param array[] $data
	 * @param string|null $sheetName
	 * @param boolean $doStandardFormatting
	 * @return Worksheet
	 */
	public function addBasicSheet(
		array $header,
		array $data,
		$sheetName = null,
		$doStandardFormatting = true
	) {
		$sheet = new Worksheet($this->excel);
		if ($sheetName) {
			$sheet->setTitle($sheetName);
		}
		$sheet->fromArray($header, null, 'A1');
		if ($data) {
			$sheet->fromArray($data, null, 'A2');
		}
		if ($doStandardFormatting) {
			$lastRow = count($data) + 1;
			$lastCol = count($header);
			$sheet->applyStandardSheetFormat(
					'A1:' . ExcelHelper::getExcelColumnName($lastCol) . $lastRow
				);
		}
		return $sheet;
	}

	/**
	 * @param string $dateFormat constants from DateFormatConstants OR custom string
	 * @return Workbook
	 */
	public function setDefaultDateFormat($dateFormat) {
		$this->dateFormat = $dateFormat;
		return $this;
	}

	/**
	 * Save Excel to given path
	 * @param string $path Full path
	 * @param string $type xls, xlsx or anything from PHPExcel_IOFactory::$_autoResolveClasses
	 * @throws PhpExcelException
	 */
	public function save($path, $type = 'xlsx') {
		$writer = $this->getWriter($type);
		try {
			$writer->save($path);
		} catch (PhpOffice_PHPExcel_Writer_Exception $e) {
			throw new PhpExcelException($e->getMessage(), 0, $e);
		}
	}

	/**
	 * Get PhpOffice PhpExcel Writer
	 * @param string $type
	 * @return PhpOffice_PHPExcel_Writer_IWriter
	 * @throws PhpExcelException
	 */
	private function getWriter($type) {
		$simpleTypes = [
			'xls' => self::PHPEXCEL_2003,
			'xlsx' => self::PHPEXCEL_2007,
		];
		$writerType = array_key_exists($type, $simpleTypes) ? $simpleTypes[$type] : $type;
		try {
			return PhpOffice_PHPExcel_IOFactory::createWriter($this->excel, $writerType);
		} catch (PhpOffice_PHPExcel_Reader_Exception $e) {
			throw new PhpExcelException("Invalid writer type '$type'.", 0, $e);
		}
	}

}
