<?php

namespace Meridius\PhpExcel;

use \PHPExcel as PhpOffice_PHPExcel;
use \PHPExcel_IOFactory as PhpOffice_PHPExcel_IOFactory;
use \PHPExcel_Writer_IWriter as PhpOffice_PHPExcel_Writer_IWriter;
use Meridius\Helpers\ExcelHelper;
use Meridius\PhpExcel\PhpExcelException;

class Workbook extends \Nette\Object {

	const PHPEXCEL_2007 = 'Excel2007';
	const PHPEXCEL_2003 = 'Excel5';

	/** @var PhpOffice_PHPExcel */
	private $excel;

	/** @var string */
	private $dateFormat;

	/**
	 * Excel with one sheet
	 * @param boolean $withoutSheets
	 */
	public function __construct($withoutSheets = true) {
		$this->excel = new PhpOffice_PHPExcel;
		if ($withoutSheets) {
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
	 * @return \Meridius\PhpExcel\Workbook
	 */
	public function setDefaultDateFormat($dateFormat) {
		$this->dateFormat = $dateFormat;
		return $this;
	}

	/**
	 * Save Excel to given path
	 * @param string $path Full path
	 * @param string $type xls|xlsx
	 */
	public function save($path, $type = 'xlsx') {
		$writer = $this->getWriter($type);
		$writer->save($path);
	}

	/**
	 * Get PhpOffice PhpExcel Writer
	 * @param string $type xls|xlsx
	 * @return PhpOffice_PHPExcel_Writer_IWriter
	 * @throws PhpExcelException
	 */
	private function getWriter($type = 'xlsx') {
		switch ($type) {
			case 'xls':
				$writerType = self::PHPEXCEL_2003;
				break;
			case 'xlsx':
				$writerType = self::PHPEXCEL_2007;
				break;
			default:
				throw new PhpExcelException("Invalid excel type '$type'.");
		}
		return PhpOffice_PHPExcel_IOFactory::createWriter($this->excel, $writerType);
	}

}
