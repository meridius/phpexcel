<?php

/**
 * @testCase
 */

namespace MeridiusTests\PhpExcel;

use Meridius\TesterExtras\AbstractIntegrationTestCase;
use Meridius\TesterExtras\Bootstrap;
use MeridiusTests\PhpExcel\ExcelEntity\TestFileExcelEntity;
use MeridiusTests\PhpExcel\Reader\TestFileReader;
use Tester;

require_once __DIR__ . '/../../../../vendor/autoload.php';

Bootstrap::setup(__DIR__ . '/../../..');
Bootstrap::createRobotLoader([
		__DIR__ . '/../files',
		__DIR__ . '/../../../../src',
	]);

class ReaderTest extends AbstractIntegrationTestCase {

	/** @var TestFileReader */
	private $reader;

	public function setUp() {
		$storage = new TempUploadStorage(__DIR__ . '/../files');
		$this->reader = new TestFileReader($storage);
	}

	public function testReaderXls() {
		$rows = $this->reader->readFile('testFile.xls');
		foreach ($rows as $row) {
			Tester\Assert::type(TestFileExcelEntity::class, $row);
		}
	}

	public function testReaderOds() {
		$rows = $this->reader->readFile('testFile.ods');
		foreach ($rows as $row) {
			Tester\Assert::type(TestFileExcelEntity::class, $row);
		}
	}

	public function testReaderXlsx() {
		$rows = $this->reader->readFile('testFile.xlsx');
		foreach ($rows as $row) {
			Tester\Assert::type(TestFileExcelEntity::class, $row);
		}
	}

}

(new ReaderTest())->run();
