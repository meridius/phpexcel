<?php

namespace services\writers;

use \services\storages\TempExcelStorage;
use \ValidationResultEntities\DfsValidation;
use \ValidationResultEntities\DfsValidation\DfsValidationResultEntity;
use Meridius\PhpExcel\Writer;
use \Meridius\PhpExcel\Workbook;

class WriterTest extends \Tester\TestCase {

	public function __construct(TempExcelStorage $tempExcelStorage) {
		parent::__construct($tempExcelStorage);
	}

	/**
	 *
	 * @param DfsValidationResultEntity $validationData
	 * @return string file path
	 */
	public function writeFile(DfsValidationResultEntity $validationData) {
		$excel = Writer::createNew();
		$this->writeSites($excel, $validationData->getSiteItemArray());
		return $this->tempExcelStorage->savePhpExcelWorkbook($excel, true);
	}

	/**
	 *
	 * @param Workbook $excel
	 * @param DfsValidation\SiteItem[] $data
	 */
	private function writeSites(Workbook $excel, array $data) {
		$header = [
			'Accurate site country',
			'Accurate site code',
			'Order number',
			'Site name',
			'Site code',
			'Site location',
		];
		$dataArray = [];
		foreach ($data as $item) {
			$dataArray[] = [
				$item->accurateSiteCountry,
				$item->accurateSiteCode,
				$item->orderNumber,
				$item->siteName,
				$item->siteCode,
				$item->siteCode,
			];
		}
		$excel->addBasicSheet($header, $dataArray, 'Sites');
	}

}
