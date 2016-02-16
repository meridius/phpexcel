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
	
	public function testMisc() {
		$a = 'A';
		Assert::same('B', ++$a);
		$af = 'AF';
		Assert::same('AG', ++$af);
		$z = 'Z';
		Assert::same('AA', ++$z);
		$ag4 = 'AG4';
		Assert::same('AG5', ++$ag4);
	}
		
	/**
	 * @dataProvider getPassingData
	 */
	public function testPassing($dimensionString, $firstCol, $firstRow, $lastCol, $lastRow, $topRightCoor, $bottomLeftCoor) {
		$dimension = new Dimension($dimensionString);
		Assert::type(Dimension::class, $dimension);
		Assert::same($firstCol, $dimension->getFirstColName());
		Assert::same($firstRow, $dimension->getFirstRowNum());
		Assert::same($lastCol, $dimension->getLastColName());
		Assert::same($lastRow, $dimension->getLastRowNum());
		Assert::same($topRightCoor, $dimension->getTopRightCoordinate());
		Assert::same($bottomLeftCoor, $dimension->getBottomLeftCoordinate());
	}
	
	public function getPassingData() {
		return [
			['A1:F4', 'A', 1, 'F', 4, 'F1', 'A4'],
			['a1:f4', 'A', 1, 'F', 4, 'F1', 'A4'],
			['AFD76:F894', 'AFD', 76, 'F', 894, 'F76', 'AFD894'],
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
