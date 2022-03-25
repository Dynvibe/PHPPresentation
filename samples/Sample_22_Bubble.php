<?php

include_once 'Sample_Header.php';

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Chart\Marker;
use PhpOffice\PhpPresentation\Shape\Chart\Series;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bubble;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Shadow;

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
            "name" => "Serie_1",
            "values" => [
                [
                    "x" => -3,
                    "y" => 10,
                    "size" => 5,
                ],
                [
                    "x" => 10,
                    "y" => 50,
                    "size" => 10,
                ],
            ],
        ],
        [
            "name" => "Serie_2",
            "values" => [
                [
                    "x" => -10,
                    "y" => 30,
                    "size" => 50,
                ],
                [
                    "x" => 0,
                    "y" => 30,
                    "size" => 20,
                ],
            ],
        ],
        [
            "name" => "Serie_3",
            "values" => [
                [
                    "x" => 20,
                    "y" => 7,
                    "size" => 20,
                ],
                [
                    "x" => 10,
                    "y" => 70,
                    "size" => 20,
                ],
            ],
        ],
    ];

    // Create a scatter chart (that should be inserted in a shape)
    echo date('H:i:s') . ' Create a scatter chart (that should be inserted in a chart shape)' . EOL;
    $lineChart = new Bubble();
    $lineChart->setIsSmooth(true);

    // find all X
    $xValues = array();
    foreach ($seriesData as $data) {
        $values = $data['values'];
        foreach ($values as $value) {
            if (!in_array($value['x'], $xValues)) {
                $xValues[] = $value['x'];
            }
        }
    }

    // add X series
    $oSeries = new Series("X", $xValues);
    $oSeries->setShowSeriesName(true);
    $oSeries->getMarker()->setSymbol(Marker::SYMBOL_CIRCLE);
    $oSeries->getMarker()->getFill()
        ->setFillType(Fill::FILL_SOLID)
        ->setStartColor(new Color('FF6F3510'))
        ->setEndColor(new Color('FF6F3510'));
    $oSeries->getMarker()->getBorder()->getColor()->setRGB('FF0000');
    $oSeries->getMarker()->setSize(10);
    $lineChart->addSeries($oSeries);

    foreach ($seriesData as $data) {
        $values = $data['values'];
        $yValues = array_fill(0, count($xValues), 0);
        $sizeValues = array_fill(0, count($xValues), 0);

        foreach ($values as $value) {
            $index = array_search($value['x'], $xValues);
            $yValues[$index] = $value['y'];
            $sizeValues[$index] = $value['size'];
        }

        $oSeries = new Series($data['name'], $yValues);
        $oSeries->setShowSeriesName(true);
        $oSeries->getMarker()->setSymbol(Marker::SYMBOL_CIRCLE);
        $oSeries->getMarker()->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->setStartColor(new Color('FF6F3510'))
            ->setEndColor(new Color('FF6F3510'));
        $oSeries->getMarker()->getBorder()->getColor()->setRGB('FF0000');
        $oSeries->getMarker()->setSize(10);
        $lineChart->addSeries($oSeries);

        $oSeries = new Series("Size " . $data['name'], $sizeValues);
        $oSeries->setShowSeriesName(true);
        $oSeries->getMarker()->setSymbol(Marker::SYMBOL_CIRCLE);
        $oSeries->getMarker()->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->setStartColor(new Color('FF6F3510'))
            ->setEndColor(new Color('FF6F3510'));
        $oSeries->getMarker()->getBorder()->getColor()->setRGB('FF0000');
        $oSeries->getMarker()->setSize(10);
        $lineChart->addSeries($oSeries);

    }

    // Create a shape (chart)
    echo date('H:i:s') . ' Create a shape (chart)' . EOL;
    $shape = $currentSlide->createChartShape();
    $shape->setName('PHPPresentation Daily Download Distribution')
        ->setResizeProportional(false)
        ->setHeight(550)
        ->setWidth(700)
        ->setOffsetX(120)
        ->setOffsetY(80)
        ->setIncludeSpreadsheet(true);
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

// Save file
echo write($objPHPPresentation, basename(__FILE__, '.php'), $writers);
if (!CLI) {
    include_once 'Sample_Footer.php';
}
