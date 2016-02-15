<?php

namespace MeridiusTests\PhpExcel;

use DateTime;
use Meridius\PhpExcel\Workbook;
use Meridius\TesterExtras\AbstractIntegrationTestCase;
use Meridius\TesterExtras\Bootstrap;
use MeridiusTests\PhpExcel\ExcelEntity\TestFileExcelEntity;
use MeridiusTests\PhpExcel\Reader\TestFileAddReader;
use MeridiusTests\PhpExcel\Reader\TestFileAddWriter;
use Tester\Assert;
use Tester\FileMock;

require_once __DIR__ . '/../../../../vendor/autoload.php';

Bootstrap::setup(__DIR__ . '/../../..');
Bootstrap::createRobotLoader([
		__DIR__ . '/../files',
		__DIR__ . '/../../../../src',
	]);

class WriterAddTest extends AbstractIntegrationTestCase {

	/** @var TestFileAddWriter */
	private $writer;

	public function setUp() {
		$this->writer = new TestFileAddWriter();
	}

	public function testAddToFile() {
		$file = FileMock::create(file_get_contents(__DIR__ . '/../files/testFile.xls'));
		$data = $this->prepareTestWriteFileData();
		$workbook = $this->writer->writeToFile($file, $data);
		Assert::type(Workbook::class, $workbook);
		$workbook->save($file, 'xls');

		$storage = new TempUploadStorage(__DIR__ . '/../files');
		$reader = new TestFileAddReader($storage);

		// new (added) tab
		$newDataRed = $reader->readAddedList($file);
		foreach ($newDataRed as $key => $rowRed) {
			Assert::equal($data[$key - 2], $rowRed); // the 2 is row where data from red excel starts
		}

		// old (existing) tab
		$oldDataRed = $reader->readOldList($file);
		foreach ($oldDataRed as $key => $rowRed) {
			Assert::equal($data[$key - 2], $rowRed); // the 2 is row where data from red excel starts
		}
	}

	private function prepareTestWriteFileData() {
		$rows = [
			[
				45.5,
				DateTime::createFromFormat("d.m.Y H:i:s", "04.09.1945 00:00:00"),
				'jrt rt g',
				5,
			], [
				4186.748,
				DateTime::createFromFormat('Y-m-d H:i:s', '2015-12-20 00:00:00'),
				'hrtjhnrt tzdrth d',
				8,
			], [
				845.48,
				DateTime::createFromFormat('d\.m\.Y H:i:s', '08.09.2005 00:00:00'),
				'ku',
				4,
			]
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

(new WriterAddTest())->run();
