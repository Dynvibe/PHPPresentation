<?php

include_once 'Sample_Header.php';

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Chart\Legend;
use PhpOffice\PhpPresentation\Shape\Chart\Marker;
use PhpOffice\PhpPresentation\Shape\Chart\Series;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bubble;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Shadow;
use PhpOffice\PhpPresentation\Style\Outline;

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
        [
            "name" => "Serie_4",
            "values" => [
                [
                    "x" => 12,
                    "y" => 22,
                    "size" => 5,
                ],
                [
                    "x" => 25,
                    "y" => 10,
                    "size" => 10,
                ],
            ],
        ],
        [
            "name" => "Serie_5",
            "values" => [
                [
                    "x" => -54,
                    "y" => 37,
                    "size" => 50,
                ],
                [
                    "x" => 0,
                    "y" => 40,
                    "size" => 10,
                ],
            ],
        ],
        [
            "name" => "Serie_6",
            "values" => [
                [
                    "x" => 20,
                    "y" => 17,
                    "size" => 25,
                ],
                [
                    "x" => 10,
                    "y" => 72,
                    "size" => 20,
                ],
            ],
        ],
        [
            "name" => "Serie_7",
            "values" => [
                [
                    "x" => -30,
                    "y" => 20,
                    "size" => 5,
                ],
                [
                    "x" => 45,
                    "y" => 45,
                    "size" => 10,
                ],
            ],
        ],
        [
            "name" => "Serie_8",
            "values" => [
                [
                    "x" => -10,
                    "y" => 40,
                    "size" => 10,
                ],
                [
                    "x" => 10,
                    "y" => 20,
                    "size" => 12,
                ],
            ],
        ],
        [
            "name" => "Serie_9",
            "values" => [
                [
                    "x" => 15,
                    "y" => 70,
                    "size" => 30,
                ],
                [
                    "x" => 25,
                    "y" => 70,
                    "size" => 20,
                ],
            ],
        ],
        [
            "name" => "Serie_10",
            "values" => [
                [
                    "x" => -30,
                    "y" => 20,
                    "size" => 15,
                ],
                [
                    "x" => 30,
                    "y" => 50,
                    "size" => 20,
                ],
            ],
        ],
        [
            "name" => "Serie_11",
            "values" => [
                [
                    "x" => 12,
                    "y" => 15,
                    "size" => 25,
                ],
                [
                    "x" => 0,
                    "y" => 10,
                    "size" => 5,
                ],
            ],
        ],
        [
            "name" => "Serie_12",
            "values" => [
                [
                    "x" => 80,
                    "y" => 70,
                    "size" => 20,
                ],
                [
                    "x" => 80,
                    "y" => 100,
                    "size" => 30,
                ],
            ],
        ],
        [
            "name" => "Serie_13",
            "values" => [
                [
                    "x" => 15,
                    "y" => 5,
                    "size" => 5,
                ],
                [
                    "x" => 15,
                    "y" => 50,
                    "size" => 20,
                ],
            ],
        ],
        [
            "name" => "Serie_14",
            "values" => [
                [
                    "x" => 27,
                    "y" => 27,
                    "size" => 27,
                ],
                [
                    "x" => -27,
                    "y" => 27,
                    "size" => 27,
                ],
            ],
        ],
        [
            "name" => "Serie_15",
            "values" => [
                [
                    "x" => 45,
                    "y" => 0,
                    "size" => 10,
                ],
                [
                    "x" => 10,
                    "y" => 70,
                    "size" => 20,
                ],
            ],
        ],
    ];

    $colors = [
        "FF32D1CD",
		"FF7CB5EC",
		"FF99c584",
		"FFffd500",
		"FFf3a344",
		"FFe35656",
		"FFEFBBFF",
		"FFA60CA6",
		"FF0457a9",
		"FFFF0081",
		"FF8D864A",
		"FF6c4111",
		"FFd9d4a5",
		"FF232323",
		"FF460446",
		"FF11520d", 
		"FFF1C40F",
		"FFA0D6B4",
		"FF36802D",
		"FF243F5B",
		"FF8D864A",
		"FF845422",
		"FF4BCC2E",
		"FFBE29EC",
		"FFFF0081",
		"FF0C457D",
		"FF003366",
		"FFFFDAB9",
		"FFDAA520",
		"FFFF6666",
		"FF66CDAA",
		"FFFFF68F",
		"FFB0E0E6",
		"FF8A2BE2",
		"FF696966",
		"FF000000",
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
    $lineChart->addSeries($oSeries);

    $indexSeries = 0;
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
        $oSeries->setShowSeriesName(false);
        $oSeries->setShowValue(false);
        $oSeries->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->setStartColor(new Color($colors[$indexSeries]))
            ->setEndColor(new Color($colors[$indexSeries]));
        $oSeries->getFill()->getStartColor()->setAlpha(50);
 
       
        $border = new Border();
        $border->setColor(new Color($colors[$indexSeries]));
        $border->setLineWidth(2);
        $oSeries->getMarker()->setBorder($border);


        $lineChart->addSeries($oSeries);

        $oSeries = new Series("Size " . $data['name'], $sizeValues);
        $lineChart->addSeries($oSeries);

        $indexSeries++;

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

    $shape->getLegend()->setPosition(Legend::POSITION_BOTTOM);
    $shape->getLegend()->getBorder()->setLineStyle(Border::LINE_NONE);

    $shape->getLegend()->setWidth(1);
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
