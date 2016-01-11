<?php

/**
 * Test: Meridius\PhpExcel\Reader
 *
 * @testCase Meridius\PhpExcel\ReaderTest
 * @package Meridius\PhpExcel
 */

namespace MeridiusTests\PhpExcel;

use Tester;
use Tester\Assert;
use Meridius\PhpExcel\Reader;

require_once __DIR__ . '/../bootstrap.php';

// TODO
class ReaderTest extends Tester\TestCase {

	/**
	 *
	 * @return mixed[] test data, expected result
	 */
	public function getSafeTrimData() {
		$date = new \DateTime;
		return [
			['as df', 'as df'],
			['as df ', 'as df'],
			['16851 ', '16851'],
			[16851, 16851],
			[$date, $date],
		];
	}

	/**
	 *
	 * @dataProvider getSafeTrimData
	 */
	public function testSafeTrim($in, $expected) {
		Assert::same($expected, StringHelper::safeTrim($in));
	}

}

$testCase = new ReaderTest;
$testCase->run();
