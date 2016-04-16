<?php

namespace MeridiusTests\PhpExcel;

use const TEMP_DIR;
use DateTime;
use Meridius\TesterExtras\AbstractIntegrationTestCase;
use Meridius\TesterExtras\Bootstrap;
use MeridiusTests\PhpExcel\ExcelEntity\TestFileExcelEntity;
use MeridiusTests\PhpExcel\Reader\TestFileWriter;
use MeridiusTests\PhpExcel\TempExcelStorage;
use Tester\Assert;

require_once __DIR__ . '/../../../../vendor/autoload.php';

Bootstrap::setup(__DIR__ . '/../../..');
Bootstrap::createRobotLoader([
		__DIR__ . '/../files',
		__DIR__ . '/../../../../src',
	]);

class WriterTest extends AbstractIntegrationTestCase {

	/** @var TestFileWriter */
	private $writer;

	public function setUp() {
		$storage = new TempExcelStorage(TEMP_DIR);
		$this->writer = new TestFileWriter($storage);
	}

	public function testWriteFile() {
		$data = $this->prepareTestWriteFileData();
		$file = $this->writer->writeFile($data);
		Assert::type('string', $file);
	}

	private function prepareTestWriteFileData() {
		$rows = [
			[
				45.5,
				DateTime::createFromFormat('d.m.Y', '04.09.1945'),
				'jrt rt g',
				5,
			], [
				4186.748,
				DateTime::createFromFormat('d.m.Y', '20.12.2015'),
				'hrtjhnrt tzdrth d',
				8,
			], [
				845.48,
				DateTime::createFromFormat('d.m.Y', '08.09.2005'),
				'ku',
				4,
			],
		];
		$data = [];
		foreach ($rows as $row) {
			$data[] = new TestFileExcelEntity([
				TestFileExcelEntity::COL_1 => $row[0],
				TestFileExcelEntity::COL_2 => $row[1],
				TestFileExcelEntity::COLUMN_3 => $row[2],
				TestFileExcelEntity::COLUMN_4 => $row[3],
			]);
		}
		return $data;
	}

}

(new WriterTest())->run();
