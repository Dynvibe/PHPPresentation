<?php

include_once 'Sample_Header.php';

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Chart\Legend;
use PhpOffice\PhpPresentation\Shape\Chart\Marker;
use PhpOffice\PhpPresentation\Shape\Chart\Series;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bubble;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Scatter;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Shadow;
use PhpOffice\PhpPresentation\Shape\Chart\Axis;

function fnSlide_Bubble(PhpPresentation $objPHPPresentation)
{
	global $oFill;
	global $oShadow;

	// Create templated slide
	echo EOL . date('H:i:s') . ' Create templated slide' . EOL;
	$currentSlide = createTemplatedSlide($objPHPPresentation); // local function

	// Generate sample data for fourth chart
	echo date('H:i:s') . ' Generate sample data for chart' . EOL;
	$seriesData = [
		[
			"name" => "X",
			"values" => [
				-3,
				-6,
				-15,
			],
		],
		[
			"name" => "Serie_1",
			"values" => [
				274,
				0,
				0,
			],
			"sizes" => [
				302,
				0,
				0,
			],
		],
		[
			"name" => "Serie_2",
			"values" => [
				0,
				245,
				0,
			],
			"sizes" => [
				0,
				300,
				0,
			],
		],
		[
			"name" => "Serie_3",
			"values" => [
				0,
				0,
				179,
			],
			"sizes" => [
				0,
				0,
				292,
			],
		],
	];

	// Create a scatter chart (that should be inserted in a shape)
	echo date('H:i:s') . ' Create a scatter chart (that should be inserted in a chart shape)' . EOL;
	$lineChart = new Bubble();
	$lineChart->setIsSmooth(true);
	$series = new Series('Downloads', $seriesData);
	$series->setShowSeriesName(true);
	$series->getMarker()->setSymbol(Marker::SYMBOL_CIRCLE);
	$series->getMarker()->getFill()
		->setFillType(Fill::FILL_SOLID)
		->setStartColor(new Color('FF6F3510'))
		->setEndColor(new Color('FF6F3510'));
	$series->getMarker()->getBorder()->getColor()->setRGB('FF0000');
	$series->getMarker()->setSize(10);
	$lineChart->addSeries($series);

	// Create a shape (chart)
	echo date('H:i:s') . ' Create a shape (chart)' . EOL;
	$shape = $currentSlide->createChartShape();
	$shape->setName('PHPPresentation Daily Download Distribution')
		->setResizeProportional(false)
		->setHeight(550)
		->setWidth(700)
		->setOffsetX(120)
		->setOffsetY(80);
	$shape->setShadow($oShadow);
	$shape->setFill($oFill);
	$shape->getBorder()->setLineStyle(Border::LINE_SINGLE);
	$shape->getTitle()->setText('PHPPresentation Daily Downloads');
	$shape->getTitle()->getFont()->setItalic(true);
	$shape->getPlotArea()->setType($lineChart);
	$shape->getView3D()->setRotationX(30);
	$shape->getView3D()->setPerspective(30);
	$shape->getLegend()->getBorder()->setLineStyle(Border::LINE_SINGLE);
	$shape->getLegend()->getFont()->setItalic(true);
}


// Create new PHPPresentation object
echo date('H:i:s') . ' Create new PHPPresentation object' . EOL;
$objPHPPresentation = new PhpPresentation();

// Set properties
echo date('H:i:s') . ' Set properties' . EOL;
$objPHPPresentation->getDocumentProperties()->setCreator('PHPOffice')
	->setLastModifiedBy('PHPPresentation Team')
	->setTitle('Sample 07 Title')
	->setSubject('Sample 07 Subject')
	->setDescription('Sample 07 Description')
	->setKeywords('office 2007 openxml libreoffice odt php')
	->setCategory('Sample Category');

// Remove first slide
echo date('H:i:s') . ' Remove first slide' . EOL;
$objPHPPresentation->removeSlideByIndex(0);

// Set Style
$oFill = new Fill();
$oFill->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFFFFFFF'));

$oShadow = new Shadow();
$oShadow->setVisible(true)->setDirection(45)->setDistance(10);

fnSlide_Bubble($objPHPPresentation);

//fnSlide_Scatter($objPHPPresentation);

// Save file
echo write($objPHPPresentation, basename(__FILE__, '.php'), $writers);
if (!CLI) {
	include_once 'Sample_Footer.php';
}
