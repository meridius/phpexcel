<?php

/**
 * @testCase
 */

namespace MeridiusTests\PhpExcel;

use Meridius\PhpExcel\Dimension;
use Meridius\PhpExcel\PhpExcelException;
use Meridius\TesterExtras\Bootstrap;
use Tester\Assert;
use Tester\TestCase;

require_once __DIR__ . '/../../../vendor/autoload.php';

Bootstrap::setup(__DIR__ . '/../..');
Bootstrap::createRobotLoader([
		__DIR__ . '/../../../src',
	]);

class DimensionTest extends TestCase {

	/**
	 * @dataProvider getPassingData
	 */
	public function testPassing($dimensionString,
		$firstCol, $firstRow, $lastCol, $lastRow,
		$topRightCoor, $bottomLeftCoor, $topLeftCoor, $bottomRightCoor
	) {
		$dimension = new Dimension($dimensionString);
		Assert::type(Dimension::class, $dimension);
		Assert::same($firstCol, $dimension->getFirstColName());
		Assert::same($firstRow, $dimension->getFirstRowNum());
		Assert::same($lastCol, $dimension->getLastColName());
		Assert::same($lastRow, $dimension->getLastRowNum());
		Assert::same($topRightCoor, (string) $dimension->getTopRightCoordinate());
		Assert::same($bottomLeftCoor, (string) $dimension->getBottomLeftCoordinate());
		Assert::same($topLeftCoor, (string) $dimension->getTopLeftCoordinate());
		Assert::same($bottomRightCoor, (string) $dimension->getBottomRightCoordinate());
	}

	public function getPassingData() {
		return [
			['A1:F4', 'A', 1, 'F', 4, 'F1', 'A4', 'A1', 'F4'],
			['a1:f4', 'A', 1, 'F', 4, 'F1', 'A4', 'A1', 'F4'],
			['AFD76:F894', 'AFD', 76, 'F', 894, 'F76', 'AFD894', 'AFD76', 'F894'],
		];
	}

	/**
	 * @dataProvider getFailingData
	 */
	public function testFailing($dimensionString) {
		Assert::exception(function() use ($dimensionString) {
			$dimension = new Dimension($dimensionString);
		}, PhpExcelException::class);
	}

	public function getFailingData() {
		return [
			[''],
			[null],
			[[]],
			[547],
			['sad'],
			['-AS54:G8'],
			['AS54:-G8'],
			['R7C5:R78C4'],
		];
	}

}

(new DimensionTest())->run();
