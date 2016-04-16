<?php

/**
 * @testCase
 */

namespace MeridiusTests\PhpExcel;

use Meridius\PhpExcel\Coordinate;
use Meridius\PhpExcel\PhpExcelException;
use Meridius\TesterExtras\Bootstrap;
use Tester\Assert;
use Tester\TestCase;

require_once __DIR__ . '/../../../vendor/autoload.php';

Bootstrap::setup(__DIR__ . '/../..');
Bootstrap::createRobotLoader([
		__DIR__ . '/../../../src',
	]);

class CoordinateTest extends TestCase {

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
	public function testPassing($coordinateString,
		$col, $row,
		$colPlusOne, $rowPlusOne, $colMinusOne, $rowMinusOne, $colMinusZero, $rowMinusZero
	) {
		$coordinate = new Coordinate($coordinateString);
		Assert::type(Coordinate::class, $coordinate);
		Assert::same(strtoupper($coordinateString), (string) $coordinate);
		Assert::same($col, $coordinate->getColName());
		Assert::same($row, $coordinate->getRowNum());
		Assert::same($colPlusOne, (string) $coordinate->shiftColBy(3));
		Assert::same($rowPlusOne, (string) $coordinate->shiftRowBy(1));
		Assert::same($colMinusOne, (string) $coordinate->shiftColBy(-1));
		Assert::same($rowMinusOne, (string) $coordinate->shiftRowBy(-1));
		Assert::same($colMinusZero, (string) $coordinate->shiftColBy());
		Assert::same($rowMinusZero, (string) $coordinate->shiftRowBy());
	}

	public function getPassingData() {
		return [
			['A1', 'A', 1, 'D1', 'D2', 'C2', 'C1', 'C1', 'C1'],
			['a1', 'A', 1, 'D1', 'D2', 'C2', 'C1', 'C1', 'C1'],
		];
	}

	public function testOutOfBounds() {
		$coordinate = new Coordinate('A1');
		Assert::exception(function() use ($coordinate) {
			$coordinate->shiftRowBy(-1);
		}, PhpExcelException::class);
		Assert::exception(function() use ($coordinate) {
			$coordinate->shiftColBy(-1);
		}, PhpExcelException::class);
	}

	/**
	 * @dataProvider getFailingData
	 */
	public function testFailing($coordinateString) {
		Assert::exception(function() use ($coordinateString) {
			$coordinate = new Coordinate($coordinateString);
		}, PhpExcelException::class);
	}

	public function getFailingData() {
		return [
			[''],
			[null],
			[[]],
			[547],
			['sad'],
			['-AS54'],
			['AS-54'],
			['-AS54:G8'],
			['AS54:-G8'],
			['R7C5:R78C4'],
		];
	}

	public function testFailingShiftBy() {
		$coordinate = new Coordinate('A1');
		Assert::exception(function() use ($coordinate) {
			$coordinate->shiftColBy('not int');
		}, PhpExcelException::class);
		Assert::exception(function() use ($coordinate) {
			$coordinate->shiftRowBy('not int');
		}, PhpExcelException::class);
	}

}

(new CoordinateTest())->run();
