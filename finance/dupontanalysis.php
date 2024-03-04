<?php

/** Error reporting */
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/London');

define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

date_default_timezone_set('Europe/London');

/**
 * PHPExcel
 *
 * Copyright (C) 2006 - 2014 PHPExcel
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPExcel
 * @package    PHPExcel
 * @copyright  Copyright (c) 2006 - 2014 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    1.8.0, 2014-03-02
 */

/** PHPExcel */
require_once 'PHPExcel_1.8.0_doc/Classes/PHPExcel.php';

//$conn=mysqli_connect("localhost","root","","finance") or die(mysqli_error());
	
$conn = mysql_connect("localhost","nacheeso_hemant","7ByV+sA[CsL6") or die(mysql_error());
mysql_select_db("nacheeso_finance");


$query1 = "select income_statement.finyear, income_statement.totalrevenue_sales,income_statement.costrevenue,
income_statement.netprofit, 
balance_sheet.totalassets, balance_sheet.inventory, balance_sheet.accountrecive, balance_sheet.totalequity from income_statement, 
balance_sheet where income_statement.company='ABC Ltd.' and income_statement.finyear=balance_sheet.finyear ";

$resultnew = mysql_query($query1, $conn);
//$resultnew = mysqli_query($conn,$query5);
$results33 = array();
$results34 = array();
$results35 = array();
$results36 = array();
$results37 = array();
$results40 = array();
$results1 = array();
$results2 = array();
$results3 = array();
$results4 = array();
$results5 = array();
$results6 = array();
$results7 = array();
$results8 = array();
$results9 = array();


//while ($info5 = mysqli_fetch_array($resultnew,MYSQLI_ASSOC))
$i=0;
$j=0;
while ($info5 = mysql_fetch_array($resultnew,MYSQL_ASSOC))
{
	$finyear = $info5['finyear'];
	$netprofit= $info5['netprofit'];
	$totalequity=$info5['totalequity'];
	
	$revsales = $info5['totalrevenue_sales'];
	$costrevenue =$info5['costrevenue'];
	$totalassets =$info5['totalassets'];
	$inventory =$info5['inventory'];
	$accountreceive=$info5['accountrecive'];

	$results33[$i][$j] = $finyear;
	$results34[$i][$j] = $finyear;
	$results35[$i][$j] = $finyear;
	$results36[$i][$j] = $finyear;
	$results37[$i][$j] = $finyear;
	$results40[$i][$j] = $finyear;
	$results1[$i][$j] = $finyear;
	$results2[$i][$j] = $finyear;
	$results3[$i][$j] = $finyear;
	$results4[$i][$j] = $finyear;
	$results5[$i][$j] = $finyear;
	$results6[$i][$j] = $finyear;
	$results7[$i][$j] = $finyear;
	$results8[$i][$j] = $finyear;
	$results9[$i][$j] = $finyear;
	
	
	$j=$j+1;	

	$results33[$i+1][$j]=$revsales/$totalassets;
	$results34[$i+1][$j]=$costrevenue/$inventory;
	$results35[$i+1][$j]=$revsales/$accountreceive;
	$results36[$i+1][$j]=round($accountreceive/($revsales/365),0);
	$results37[$i+1][$j]=round($inventory/($costrevenue/365),0);
	$results40[$i+1][$j]=$netprofit/$totalassets;
	$results1[$i+1][$j]=$netprofit/$totalequity;
	$results2[$i+1][$j]=$netprofit/$totalassets;
	$results3[$i+1][$j]=$totalassets/$totalequity;
	$results4[$i+1][$j]= $netprofit/$revsales;
	$results5[$i+1][$j]= $revsales/$totalassets;
	$results6[$i+1][$j]= $revsales;
	$results7[$i+1][$j]= $netprofit;
	$results8[$i+1][$j]= $revsales;
	$results9[$i+1][$j]= $totalassets;

	
	}
 
$objPHPExcel = new PHPExcel();
$objWorksheet = $objPHPExcel->getActiveSheet();

$sheet = $objPHPExcel->getActiveSheet();
$sheet->getStyle('A1:T1')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->getStyle('A1:T1')->getFill()->getStartColor()->setRGB('6a5acd');

$objPHPExcel->getActiveSheet()->mergeCells('A1:T1');
$objPHPExcel->getActiveSheet()->setCellValue('A1', 'Du Pont Dashboard');
$objPHPExcel->getActiveSheet()->getStyle("A1")->getFont()->setBold(true)
                                ->setName('Verdana')
                                ->setSize(20)
								->getColor()->setRGB('FFFFF');

$objPHPExcel->getActiveSheet()->getStyle('A1:T1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
								
$objWorksheet->fromArray(
$results1,NULL,'I8'
);


$dataseriesLabels1 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->mergeCells('I5:I7');
$objPHPExcel->getActiveSheet()->setCellValue('I5', 'Return on Equity');
$objPHPExcel->getActiveSheet()->getStyle('I5')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle("I5")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(12)
								->getColor()->setRGB('6a5acd');
$objPHPExcel->getActiveSheet()->getStyle('I5:I7')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
								
$styleArray = array(
       'borders' => array(
             'outline' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('argb' => '0000FF'),
             ),
       ),
);

$objPHPExcel->getActiveSheet()->getStyle('I5:N16')->applyFromArray($styleArray);
												
														
                  
$xAxisTickValues1 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$I$8:$M8', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues1 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$I$9:$M9', NULL, 5));

//	Build the dataseries


$series1= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues1)-1),			// plotOrder
	$dataseriesLabels1,							// plotLabel
	$xAxisTickValues1,								// plotCategory
	$dataSeriesValues1								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series1->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

$layout1 = new PHPExcel_Chart_Layout();

//	Set the series in the plot area
$plotarea1 = new PHPExcel_Chart_PlotArea($layout1, array($series1));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend1=NULL;
$title1=new PHPExcel_Chart_Title('');
$yAxisLabel1=NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P9:P10');
$objPHPExcel->getActiveSheet()->mergeCells('R9:R10');
$objPHPExcel->getActiveSheet()->setCellValue('J5', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('L5', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('J6', '=IF(M9>I9,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('L6', '=IF(M9>L9,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('L7', '=round((M9-I9)/ABS(I9),2)');
$objPHPExcel->getActiveSheet()->setCellValue('J7', '=round((M9-L9)/ABS(L9),2)');
$objPHPExcel->getActiveSheet()->getStyle('J6')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('L6')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

//conditions
$objConditional1 = new PHPExcel_Style_Conditional();
$objConditional1->setConditionType(PHPExcel_Style_Conditional::CONDITION_CELLIS)
    ->setOperatorType(PHPExcel_Style_Conditional::OPERATOR_LESSTHAN)
    ->addCondition('0');
$objConditional1->getStyle()->getFill()->applyFromArray(
            array(
            'type'       => PHPExcel_Style_Fill::FILL_SOLID,
            'startcolor' => array('argb' => 'FF0000'),
			'endcolor' => array('argb' => 'FF0000')
	));
$objConditional1->getStyle()->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_WHITE);
$objConditional1->getStyle()->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE);


$objConditional3 = new PHPExcel_Style_Conditional();
$objConditional3->setConditionType(PHPExcel_Style_Conditional::CONDITION_CELLIS)
    ->setOperatorType(PHPExcel_Style_Conditional::OPERATOR_GREATERTHANOREQUAL)
    ->addCondition('0');
$objConditional3->getStyle()->getFill()->applyFromArray(
            array(
            'type'       => PHPExcel_Style_Fill::FILL_SOLID,
            'startcolor' => array('argb'=>'008000'),
            'endcolor' => array('argb'=>'008000')
	 ));
$objConditional3->getStyle()->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_WHITE);
$objConditional3->getStyle()->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE);
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('J7')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('L7')->setConditionalStyles($conditionalStyles);
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('L7')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);


$objPHPExcel->getActiveSheet()->getStyle('J7')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('M7')->setConditionalStyles($conditionalStyles);


//conditions

//	Create the chart
$chart1= new PHPExcel_Chart(
	'chart1',		// name
	$title1,		// title
	$legend1,		// legend
	$plotarea1,	// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel1	// yAxisLabel
);

//	Set the position where the chart should appear in the worksheet

$chart1->setTopLeftPosition('I10');
$chart1->setBottomRightPosition('N16');

//	Add the chart to the worksheet
$objWorksheet->addChart($chart1);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////


$objWorksheet->fromArray(
$results2,NULL,'D23'
);

$dataseriesLabels2 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->mergeCells('D20:D22');
$objPHPExcel->getActiveSheet()->setCellValue('D20', 'Return on Assets');
$objPHPExcel->getActiveSheet()->getStyle('D20')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle("D20")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(12)
								->getColor()->setRGB('6a5acd');
$objPHPExcel->getActiveSheet()->getStyle('D20:D22')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
								
$styleArray = array(
       'borders' => array(
             'outline' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('argb' => '0000FF'),
             ),
       ),
);

$objPHPExcel->getActiveSheet()->getStyle('D20:I31')->applyFromArray($styleArray);
												
														
                  
$xAxisTickValues2 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$D$23:$H23', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues2 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$D$24:$H24', NULL, 5));

//	Build the dataseries


$series2= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues2)-1),			// plotOrder
	$dataseriesLabels2,							// plotLabel
	$xAxisTickValues2,								// plotCategory
	$dataSeriesValues2								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series2->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

$layout2 = new PHPExcel_Chart_Layout();


//	Set the series in the plot area
$plotarea2 = new PHPExcel_Chart_PlotArea($layout2,array($series2));

//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend2=NULL;
$title2=new PHPExcel_Chart_Title('');
$yAxisLabel2=NULL;

$objPHPExcel->getActiveSheet()->mergeCells('E22:E22');
$objPHPExcel->getActiveSheet()->mergeCells('G22:G22');
$objPHPExcel->getActiveSheet()->setCellValue('E20', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('G20', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('E21', '=IF(H24>D24,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('G21', '=IF(H24>G24,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('E22', '=round((H24-D24)/ABS(D24),2)');
$objPHPExcel->getActiveSheet()->setCellValue('G22', '=round((H24-G24)/ABS(G24),2)');
$objPHPExcel->getActiveSheet()->getStyle('E21')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('G21')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

//conditions
$objConditional1 = new PHPExcel_Style_Conditional();
$objConditional1->setConditionType(PHPExcel_Style_Conditional::CONDITION_CELLIS)
    ->setOperatorType(PHPExcel_Style_Conditional::OPERATOR_LESSTHAN)
    ->addCondition('0');
$objConditional1->getStyle()->getFill()->applyFromArray(
            array(
            'type'       => PHPExcel_Style_Fill::FILL_SOLID,
            'startcolor' => array('argb' => 'FF0000'),
			'endcolor' => array('argb' => 'FF0000')
	));
$objConditional1->getStyle()->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_WHITE);
$objConditional1->getStyle()->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE);


$objConditional3 = new PHPExcel_Style_Conditional();
$objConditional3->setConditionType(PHPExcel_Style_Conditional::CONDITION_CELLIS)
    ->setOperatorType(PHPExcel_Style_Conditional::OPERATOR_GREATERTHANOREQUAL)
    ->addCondition('0');
$objConditional3->getStyle()->getFill()->applyFromArray(
            array(
            'type'       => PHPExcel_Style_Fill::FILL_SOLID,
            'startcolor' => array('argb'=>'008000'),
            'endcolor' => array('argb'=>'008000')
	 ));
	 
$objConditional3->getStyle()->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_WHITE);
$objConditional3->getStyle()->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE);

$objPHPExcel->getActiveSheet()->getStyle('E22')->setConditionalStyles($conditionalStyles);
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('E22')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);

$objPHPExcel->getActiveSheet()->getStyle('G22')->setConditionalStyles($conditionalStyles);
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('G22')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);


$objPHPExcel->getActiveSheet()->getStyle('E22')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('G22')->setConditionalStyles($conditionalStyles);


//conditions





//	Create the chart
$chart2= new PHPExcel_Chart(
	'chart2',		// name
	$title2,		// title
	$legend2,		// legend
	$plotarea2,	// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel2	// yAxisLabel
);

//	Set the position where the chart should appear in the worksheet

$chart2->setTopLeftPosition('D25');
$chart2->setBottomRightPosition('I31');

//	Add the chart to the worksheet
$objWorksheet->addChart($chart2);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////


$objWorksheet->fromArray(
$results3,NULL,'N23'
);

$dataseriesLabels3 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->mergeCells('N20:N22');
$objPHPExcel->getActiveSheet()->setCellValue('N20', 'Return of Equity');
$objPHPExcel->getActiveSheet()->getStyle('N20')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle("N20")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(12)
								->getColor()->setRGB('6a5acd');
$objPHPExcel->getActiveSheet()->getStyle('N20:R22')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
								
								
$styleArray = array(
       'borders' => array(
             'outline' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('argb' => '0000FF'),
             ),
       ),
);

$objPHPExcel->getActiveSheet()->getStyle('N20:S31')->applyFromArray($styleArray);
												
														
                  
$xAxisTickValues3 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$N$23:$R23', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues3 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$N$24:$R24', NULL, 5));

//	Build the dataseries


$series3= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues3)-1),			// plotOrder
	$dataseriesLabels3,							// plotLabel
	$xAxisTickValues3,								// plotCategory
	$dataSeriesValues3								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series3->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

$layout3 = new PHPExcel_Chart_Layout();


//	Set the series in the plot area
$plotarea3 = new PHPExcel_Chart_PlotArea($layout3, array($series3));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend3=NULL;
$title3=new PHPExcel_Chart_Title('');   
$yAxisLabel3=NULL;


$objPHPExcel->getActiveSheet()->mergeCells('O22:O22');
$objPHPExcel->getActiveSheet()->mergeCells('Q22:Q22');
$objPHPExcel->getActiveSheet()->setCellValue('O20', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('Q20', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('O21', '=IF(R24>N24,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('Q21', '=IF(R24>N24,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('O22', '=round((R24-N24)/ABS(N24),2)');
$objPHPExcel->getActiveSheet()->setCellValue('Q22', '=round((R24-Q24)/ABS(Q24),2)');
$objPHPExcel->getActiveSheet()->getStyle('O21')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('Q21')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

//conditions
$objConditional1 = new PHPExcel_Style_Conditional();
$objConditional1->setConditionType(PHPExcel_Style_Conditional::CONDITION_CELLIS)
    ->setOperatorType(PHPExcel_Style_Conditional::OPERATOR_LESSTHAN)
    ->addCondition('0');
$objConditional1->getStyle()->getFill()->applyFromArray(
            array(
            'type'       => PHPExcel_Style_Fill::FILL_SOLID,
            'startcolor' => array('argb' => 'FF0000'),
			'endcolor' => array('argb' => 'FF0000')
	));
$objConditional1->getStyle()->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_WHITE);
$objConditional1->getStyle()->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE);


$objConditional3 = new PHPExcel_Style_Conditional();
$objConditional3->setConditionType(PHPExcel_Style_Conditional::CONDITION_CELLIS)
    ->setOperatorType(PHPExcel_Style_Conditional::OPERATOR_GREATERTHANOREQUAL)
    ->addCondition('0');
$objConditional3->getStyle()->getFill()->applyFromArray(
            array(
            'type'       => PHPExcel_Style_Fill::FILL_SOLID,
            'startcolor' => array('argb'=>'008000'),
            'endcolor' => array('argb'=>'008000')
	 ));
	 
$objConditional3->getStyle()->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_WHITE);
$objConditional3->getStyle()->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE);

$objPHPExcel->getActiveSheet()->getStyle('O22')->setConditionalStyles($conditionalStyles);
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('O22')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);

$objPHPExcel->getActiveSheet()->getStyle('Q22')->setConditionalStyles($conditionalStyles);
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('Q22')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);


$objPHPExcel->getActiveSheet()->getStyle('O22')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('Q22')->setConditionalStyles($conditionalStyles);


//conditions




//	Create the chart
$chart3= new PHPExcel_Chart(
	'chart3',		// name
	$title3,		// title
	$legend3,		// legend
	$plotarea3,	// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel3	// yAxisLabel
);

//	Set the position where the chart should appear in the worksheet

$chart3->setTopLeftPosition('N25');
$chart3->setBottomRightPosition('S31');

//	Add the chart to the worksheet
$objWorksheet->addChart($chart3);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////



$objWorksheet->fromArray(
$results4,NULL,'A38'
);

$dataseriesLabels4 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->mergeCells('A35:A37');
$objPHPExcel->getActiveSheet()->setCellValue('A35', 'Net Profit Margin');
$objPHPExcel->getActiveSheet()->getStyle('A35')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle("A35")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(12)
								->getColor()->setRGB('6a5acd');
$objPHPExcel->getActiveSheet()->getStyle('A35:A37')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
								
								
$styleArray = array(
       'borders' => array(
             'outline' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('argb' => '0000FF'),
             ),
       ),
);

$objPHPExcel->getActiveSheet()->getStyle('A35:F46')->applyFromArray($styleArray);
												
														
                  
$xAxisTickValues4 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$38:$E38', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues4 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$A$39:$E39', NULL, 5));

//	Build the dataseries


$series4= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues4)-1),			// plotOrder
	$dataseriesLabels4,							// plotLabel
	$xAxisTickValues4,								// plotCategory
	$dataSeriesValues4								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series4->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

$layout4 = new PHPExcel_Chart_Layout();


//	Set the series in the plot area
$plotarea4 = new PHPExcel_Chart_PlotArea($layout4, array($series4));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend4=NULL;
$title4=new PHPExcel_Chart_Title('');   
$yAxisLabel4=NULL;


$objPHPExcel->getActiveSheet()->mergeCells('B37:B37');
$objPHPExcel->getActiveSheet()->mergeCells('D37:D37');
$objPHPExcel->getActiveSheet()->setCellValue('B35', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('D35', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('B36', '=IF(E39>A39,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('D36', '=IF(E39>D39,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('B37', '=round((E39-A39)/ABS(A39),2)');
$objPHPExcel->getActiveSheet()->setCellValue('D37', '=round((E39-D39)/ABS(D39),2)');
$objPHPExcel->getActiveSheet()->getStyle('B36')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('D36')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

//conditions
$objConditional1 = new PHPExcel_Style_Conditional();
$objConditional1->setConditionType(PHPExcel_Style_Conditional::CONDITION_CELLIS)
    ->setOperatorType(PHPExcel_Style_Conditional::OPERATOR_LESSTHAN)
    ->addCondition('0');
$objConditional1->getStyle()->getFill()->applyFromArray(
            array(
            'type'       => PHPExcel_Style_Fill::FILL_SOLID,
            'startcolor' => array('argb' => 'FF0000'),
			'endcolor' => array('argb' => 'FF0000')
	));
$objConditional1->getStyle()->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_WHITE);
$objConditional1->getStyle()->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE);


$objConditional3 = new PHPExcel_Style_Conditional();
$objConditional3->setConditionType(PHPExcel_Style_Conditional::CONDITION_CELLIS)
    ->setOperatorType(PHPExcel_Style_Conditional::OPERATOR_GREATERTHANOREQUAL)
    ->addCondition('0');
$objConditional3->getStyle()->getFill()->applyFromArray(
            array(
            'type'       => PHPExcel_Style_Fill::FILL_SOLID,
            'startcolor' => array('argb'=>'008000'),
            'endcolor' => array('argb'=>'008000')
	 ));
	 
$objConditional3->getStyle()->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_WHITE);
$objConditional3->getStyle()->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE);

$objPHPExcel->getActiveSheet()->getStyle('B37')->setConditionalStyles($conditionalStyles);
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('B37')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);

$objPHPExcel->getActiveSheet()->getStyle('D37')->setConditionalStyles($conditionalStyles);
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('D37')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);


$objPHPExcel->getActiveSheet()->getStyle('B37')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('D37')->setConditionalStyles($conditionalStyles);


//conditions




//	Create the chart
$chart4= new PHPExcel_Chart(
	'chart4',		// name
	$title4,		// title
	$legend4,		// legend
	$plotarea4,	// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel4	// yAxisLabel
);

//	Set the position where the chart should appear in the worksheet

$chart4->setTopLeftPosition('A40');
$chart4->setBottomRightPosition('F46');

//	Add the chart to the worksheet
$objWorksheet->addChart($chart4);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////


$objWorksheet->fromArray(
$results5,NULL,'R38'
);

$dataseriesLabels5 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->mergeCells('R35:R37');
$objPHPExcel->getActiveSheet()->setCellValue('R35', 'Total Assets Turonver');
$objPHPExcel->getActiveSheet()->getStyle('R35')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle("R35")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(12)
								->getColor()->setRGB('6a5acd');
$objPHPExcel->getActiveSheet()->getStyle('R35:R37')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
								
								
$styleArray = array(
       'borders' => array(
             'outline' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('argb' => '0000FF'),
             ),
       ),
);

$objPHPExcel->getActiveSheet()->getStyle('R35:W46')->applyFromArray($styleArray);
												
														
                  
$xAxisTickValues5 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$R$38:$V38', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues5 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$R$39:$V39', NULL, 5));

//	Build the dataseries


$series5= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues5)-1),			// plotOrder
	$dataseriesLabels5,							// plotLabel
	$xAxisTickValues5,								// plotCategory
	$dataSeriesValues5								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series5->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

$layout5 = new PHPExcel_Chart_Layout();


//	Set the series in the plot area
$plotarea5 = new PHPExcel_Chart_PlotArea($layout5, array($series5));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend5=NULL;
$title5=new PHPExcel_Chart_Title('');   
$yAxisLabel5=NULL;


$objPHPExcel->getActiveSheet()->mergeCells('S37:S37');
$objPHPExcel->getActiveSheet()->mergeCells('U37:U37');
$objPHPExcel->getActiveSheet()->setCellValue('S35', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('U35', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('S36', '=IF(V39>R39,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('U36', '=IF(V39>U39,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('S37', '=round((V39-R39)/ABS(R39),2)');
$objPHPExcel->getActiveSheet()->setCellValue('U37', '=round((V39-U39)/ABS(U39),2)');
$objPHPExcel->getActiveSheet()->getStyle('S36')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('U36')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

//conditions
$objConditional1 = new PHPExcel_Style_Conditional();
$objConditional1->setConditionType(PHPExcel_Style_Conditional::CONDITION_CELLIS)
    ->setOperatorType(PHPExcel_Style_Conditional::OPERATOR_LESSTHAN)
    ->addCondition('0');
$objConditional1->getStyle()->getFill()->applyFromArray(
            array(
            'type'       => PHPExcel_Style_Fill::FILL_SOLID,
            'startcolor' => array('argb' => 'FF0000'),
			'endcolor' => array('argb' => 'FF0000')
	));
$objConditional1->getStyle()->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_WHITE);
$objConditional1->getStyle()->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE);


$objConditional3 = new PHPExcel_Style_Conditional();
$objConditional3->setConditionType(PHPExcel_Style_Conditional::CONDITION_CELLIS)
    ->setOperatorType(PHPExcel_Style_Conditional::OPERATOR_GREATERTHANOREQUAL)
    ->addCondition('0');
$objConditional3->getStyle()->getFill()->applyFromArray(
            array(
            'type'       => PHPExcel_Style_Fill::FILL_SOLID,
            'startcolor' => array('argb'=>'008000'),
            'endcolor' => array('argb'=>'008000')
	 ));
	 
$objConditional3->getStyle()->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_WHITE);
$objConditional3->getStyle()->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE);

$objPHPExcel->getActiveSheet()->getStyle('S37')->setConditionalStyles($conditionalStyles);
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('S37')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);

$objPHPExcel->getActiveSheet()->getStyle('U37')->setConditionalStyles($conditionalStyles);
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('U37')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);


$objPHPExcel->getActiveSheet()->getStyle('S37')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('U37')->setConditionalStyles($conditionalStyles);


//conditions




//	Create the chart
$chart5= new PHPExcel_Chart(
	'chart5',		// name
	$title5,		// title
	$legend5,		// legend
	$plotarea5,	// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel5	// yAxisLabel
);

//	Set the position where the chart should appear in the worksheet

$chart5->setTopLeftPosition('R40');
$chart5->setBottomRightPosition('W46');

//	Add the chart to the worksheet
$objWorksheet->addChart($chart5);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////




$objWorksheet->fromArray(
$results6,NULL,'A54'
);

$dataseriesLabels6 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->mergeCells('A51:A53');
$objPHPExcel->getActiveSheet()->setCellValue('A51', 'Sales');
$objPHPExcel->getActiveSheet()->getStyle('A51')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle("A51")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(12)
								->getColor()->setRGB('6a5acd');
$objPHPExcel->getActiveSheet()->getStyle('A51:A51')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
								
								
$styleArray = array(
       'borders' => array(
             'outline' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('argb' => '0000FF'),
             ),
       ),
);

$objPHPExcel->getActiveSheet()->getStyle('A51:F62')->applyFromArray($styleArray);
												
														
                  
$xAxisTickValues6 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$54:$E54', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues6 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$A$55:$E55', NULL, 5));

//	Build the dataseries


$series6= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues6)-1),			// plotOrder
	$dataseriesLabels6,							// plotLabel
	$xAxisTickValues6,								// plotCategory
	$dataSeriesValues6								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series6->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

$layout6 = new PHPExcel_Chart_Layout();


//	Set the series in the plot area
$plotarea6 = new PHPExcel_Chart_PlotArea($layout6, array($series6));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend6=NULL;
$title6=new PHPExcel_Chart_Title('');   
$yAxisLabel6=NULL;


$objPHPExcel->getActiveSheet()->mergeCells('B53:B53');
$objPHPExcel->getActiveSheet()->mergeCells('D53:D53');
$objPHPExcel->getActiveSheet()->setCellValue('B51', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('D51', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('B52', '=IF(E55>A55,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('D52', '=IF(E55>D55,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('B53', '=round((E55-A55)/ABS(A55),2)');
$objPHPExcel->getActiveSheet()->setCellValue('D53', '=round((E55-D55)/ABS(D55),2)');
$objPHPExcel->getActiveSheet()->getStyle('B52')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('D52')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

//conditions
$objConditional1 = new PHPExcel_Style_Conditional();
$objConditional1->setConditionType(PHPExcel_Style_Conditional::CONDITION_CELLIS)
    ->setOperatorType(PHPExcel_Style_Conditional::OPERATOR_LESSTHAN)
    ->addCondition('0');
$objConditional1->getStyle()->getFill()->applyFromArray(
            array(
            'type'       => PHPExcel_Style_Fill::FILL_SOLID,
            'startcolor' => array('argb' => 'FF0000'),
			'endcolor' => array('argb' => 'FF0000')
	));
$objConditional1->getStyle()->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_WHITE);
$objConditional1->getStyle()->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE);


$objConditional3 = new PHPExcel_Style_Conditional();
$objConditional3->setConditionType(PHPExcel_Style_Conditional::CONDITION_CELLIS)
    ->setOperatorType(PHPExcel_Style_Conditional::OPERATOR_GREATERTHANOREQUAL)
    ->addCondition('0');
$objConditional3->getStyle()->getFill()->applyFromArray(
            array(
            'type'       => PHPExcel_Style_Fill::FILL_SOLID,
            'startcolor' => array('argb'=>'008000'),
            'endcolor' => array('argb'=>'008000')
	 ));
	 
$objConditional3->getStyle()->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_WHITE);
$objConditional3->getStyle()->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE);

$objPHPExcel->getActiveSheet()->getStyle('B53')->setConditionalStyles($conditionalStyles);
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('B53')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);

$objPHPExcel->getActiveSheet()->getStyle('D53')->setConditionalStyles($conditionalStyles);
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('D53')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);


$objPHPExcel->getActiveSheet()->getStyle('B53')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('D53')->setConditionalStyles($conditionalStyles);


//conditions




//	Create the chart
$chart6= new PHPExcel_Chart(
	'chart6',		// name
	$title6,		// title
	$legend6,		// legend
	$plotarea6,	// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel6	// yAxisLabel
);

//	Set the position where the chart should appear in the worksheet

$chart6->setTopLeftPosition('A56');
$chart6->setBottomRightPosition('F62');

//	Add the chart to the worksheet
$objWorksheet->addChart($chart6);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////





$objWorksheet->fromArray(
$results7,NULL,'H54'
);

$dataseriesLabels7 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->mergeCells('H51:H53');
$objPHPExcel->getActiveSheet()->setCellValue('H51', 'Net Profit');
$objPHPExcel->getActiveSheet()->getStyle('H51')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle("H51")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(12)
								->getColor()->setRGB('6a5acd');
$objPHPExcel->getActiveSheet()->getStyle('H51:H53')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
								
								
$styleArray = array(
       'borders' => array(
             'outline' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('argb' => '0000FF'),
             ),
       ),
);

$objPHPExcel->getActiveSheet()->getStyle('H51:M62')->applyFromArray($styleArray);
												
														
                  
$xAxisTickValues7 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$H$54:$L54', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues7 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$H$55:$L55', NULL, 5));

//	Build the dataseries


$series7= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues7)-1),			// plotOrder
	$dataseriesLabels7,							// plotLabel
	$xAxisTickValues7,								// plotCategory
	$dataSeriesValues7								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series7->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

$layout7 = new PHPExcel_Chart_Layout();


//	Set the series in the plot area
$plotarea7 = new PHPExcel_Chart_PlotArea($layout7, array($series6));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend7=NULL;
$title7=new PHPExcel_Chart_Title('');   
$yAxisLabel7=NULL;


$objPHPExcel->getActiveSheet()->mergeCells('I53:I53');
$objPHPExcel->getActiveSheet()->mergeCells('K53:K53');
$objPHPExcel->getActiveSheet()->setCellValue('I51', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('K51', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('I52', '=IF(L55>H55,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('K52', '=IF(L55>K55,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('I53', '=round((L55-H55)/ABS(H55),2)');
$objPHPExcel->getActiveSheet()->setCellValue('K53', '=round((L55-K55)/ABS(K55),2)');
$objPHPExcel->getActiveSheet()->getStyle('I52')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('K52')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

//conditions
$objConditional1 = new PHPExcel_Style_Conditional();
$objConditional1->setConditionType(PHPExcel_Style_Conditional::CONDITION_CELLIS)
    ->setOperatorType(PHPExcel_Style_Conditional::OPERATOR_LESSTHAN)
    ->addCondition('0');
$objConditional1->getStyle()->getFill()->applyFromArray(
            array(
            'type'       => PHPExcel_Style_Fill::FILL_SOLID,
            'startcolor' => array('argb' => 'FF0000'),
			'endcolor' => array('argb' => 'FF0000')
	));
$objConditional1->getStyle()->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_WHITE);
$objConditional1->getStyle()->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE);


$objConditional3 = new PHPExcel_Style_Conditional();
$objConditional3->setConditionType(PHPExcel_Style_Conditional::CONDITION_CELLIS)
    ->setOperatorType(PHPExcel_Style_Conditional::OPERATOR_GREATERTHANOREQUAL)
    ->addCondition('0');
$objConditional3->getStyle()->getFill()->applyFromArray(
            array(
            'type'       => PHPExcel_Style_Fill::FILL_SOLID,
            'startcolor' => array('argb'=>'008000'),
            'endcolor' => array('argb'=>'008000')
	 ));
	 
$objConditional3->getStyle()->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_WHITE);
$objConditional3->getStyle()->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE);

$objPHPExcel->getActiveSheet()->getStyle('I53')->setConditionalStyles($conditionalStyles);
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('I53')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);

$objPHPExcel->getActiveSheet()->getStyle('K53')->setConditionalStyles($conditionalStyles);
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('K53')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);


$objPHPExcel->getActiveSheet()->getStyle('I53')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('K53')->setConditionalStyles($conditionalStyles);


//conditions




//	Create the chart
$chart7= new PHPExcel_Chart(
	'chart7',		// name
	$title7,		// title
	$legend7,		// legend
	$plotarea7,	// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel7	// yAxisLabel
);

//	Set the position where the chart should appear in the worksheet

$chart7->setTopLeftPosition('H56');
$chart7->setBottomRightPosition('M62');

//	Add the chart to the worksheet
$objWorksheet->addChart($chart7);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////



$objWorksheet->fromArray(
$results8,NULL,'O54'
);

$dataseriesLabels8 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->mergeCells('O51:O53');
$objPHPExcel->getActiveSheet()->setCellValue('O51', 'Sales');
$objPHPExcel->getActiveSheet()->getStyle('O51')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle("O51")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(12)
								->getColor()->setRGB('6a5acd');
$objPHPExcel->getActiveSheet()->getStyle('O51:O51')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
								
								
$styleArray = array(
       'borders' => array(
             'outline' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('argb' => '0000FF'),
             ),
       ),
);

$objPHPExcel->getActiveSheet()->getStyle('O51:T62')->applyFromArray($styleArray);
												
														
                  
$xAxisTickValues8 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$O$54:$S54', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues8 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$O$55:$S55', NULL, 5));

//	Build the dataseries


$series8= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues8)-1),			// plotOrder
	$dataseriesLabels8,							// plotLabel
	$xAxisTickValues8,								// plotCategory
	$dataSeriesValues8								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series8->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

$layout8 = new PHPExcel_Chart_Layout();


//	Set the series in the plot area
$plotarea8 = new PHPExcel_Chart_PlotArea($layout8, array($series6));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend8=NULL;
$title8=new PHPExcel_Chart_Title('');   
$yAxisLabel8=NULL;


$objPHPExcel->getActiveSheet()->mergeCells('P53:P53');
$objPHPExcel->getActiveSheet()->mergeCells('R53:R53');
$objPHPExcel->getActiveSheet()->setCellValue('P51', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R51', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P52', '=IF(S55>O55,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R52', '=IF(S55>R55,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P53', '=round((S55-O55)/ABS(O55),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R53', '=round((S55-R55)/ABS(R55),2)');
$objPHPExcel->getActiveSheet()->getStyle('P52')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R52')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

//conditions
$objConditional1 = new PHPExcel_Style_Conditional();
$objConditional1->setConditionType(PHPExcel_Style_Conditional::CONDITION_CELLIS)
    ->setOperatorType(PHPExcel_Style_Conditional::OPERATOR_LESSTHAN)
    ->addCondition('0');
$objConditional1->getStyle()->getFill()->applyFromArray(
            array(
            'type'       => PHPExcel_Style_Fill::FILL_SOLID,
            'startcolor' => array('argb' => 'FF0000'),
			'endcolor' => array('argb' => 'FF0000')
	));
$objConditional1->getStyle()->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_WHITE);
$objConditional1->getStyle()->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE);


$objConditional3 = new PHPExcel_Style_Conditional();
$objConditional3->setConditionType(PHPExcel_Style_Conditional::CONDITION_CELLIS)
    ->setOperatorType(PHPExcel_Style_Conditional::OPERATOR_GREATERTHANOREQUAL)
    ->addCondition('0');
$objConditional3->getStyle()->getFill()->applyFromArray(
            array(
            'type'       => PHPExcel_Style_Fill::FILL_SOLID,
            'startcolor' => array('argb'=>'008000'),
            'endcolor' => array('argb'=>'008000')
	 ));
	 
$objConditional3->getStyle()->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_WHITE);
$objConditional3->getStyle()->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE);

$objPHPExcel->getActiveSheet()->getStyle('P53')->setConditionalStyles($conditionalStyles);
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P53')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);

$objPHPExcel->getActiveSheet()->getStyle('R53')->setConditionalStyles($conditionalStyles);
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('R53')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);


$objPHPExcel->getActiveSheet()->getStyle('P53')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R53')->setConditionalStyles($conditionalStyles);


//conditions




//	Create the chart
$chart8= new PHPExcel_Chart(
	'chart8',		// name
	$title8,		// title
	$legend8,		// legend
	$plotarea8,	// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel8	// yAxisLabel
);

//	Set the position where the chart should appear in the worksheet

$chart8->setTopLeftPosition('O56');
$chart8->setBottomRightPosition('T62');

//	Add the chart to the worksheet
$objWorksheet->addChart($chart8);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////



$objWorksheet->fromArray(
$results9,NULL,'V54'
);

$dataseriesLabels9 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->mergeCells('V51:V53');
$objPHPExcel->getActiveSheet()->setCellValue('V51', 'Total Assets');
$objPHPExcel->getActiveSheet()->getStyle('V51')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle("V51")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(12)
								->getColor()->setRGB('6a5acd');
$objPHPExcel->getActiveSheet()->getStyle('V51:V51')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
								
								
$styleArray = array(
       'borders' => array(
             'outline' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('argb' => '0000FF'),
             ),
       ),
);

$objPHPExcel->getActiveSheet()->getStyle('V51:AA62')->applyFromArray($styleArray);


$styleArray1 = array(
      'borders' => array(
        'bottom' => array(
          'style' => PHPExcel_Style_Border::BORDER_THIN,
          'color' => array('argb' => '0000FF'),
        ), 
      ),
    );

$styleArray2 = array(
      'borders' => array(
        'right' => array(
          'style' => PHPExcel_Style_Border::BORDER_THIN,
          'color' => array('argb' => '0000FF'),
        ), 
      ),
    );
	
$objPHPExcel->getActiveSheet()->getStyle('G18:P18')->applyFromArray($styleArray1);
$objPHPExcel->getActiveSheet()->getStyle('F19')->applyFromArray($styleArray2);
$objPHPExcel->getActiveSheet()->getStyle('P19')->applyFromArray($styleArray2);
$objPHPExcel->getActiveSheet()->getStyle('K17:K18')->applyFromArray($styleArray2);

$objPHPExcel->getActiveSheet()->getStyle('C33:T33')->applyFromArray($styleArray1);
$objPHPExcel->getActiveSheet()->getStyle('F32:F33')->applyFromArray($styleArray2);
$objPHPExcel->getActiveSheet()->getStyle('B34')->applyFromArray($styleArray2);
$objPHPExcel->getActiveSheet()->getStyle('T34')->applyFromArray($styleArray2);


$objPHPExcel->getActiveSheet()->getStyle('R48:X48')->applyFromArray($styleArray1);
$objPHPExcel->getActiveSheet()->getStyle('Q49:Q50')->applyFromArray($styleArray2);
$objPHPExcel->getActiveSheet()->getStyle('X49:X50')->applyFromArray($styleArray2);

$objPHPExcel->getActiveSheet()->getStyle('D48:J48')->applyFromArray($styleArray1);
$objPHPExcel->getActiveSheet()->getStyle('C49:C50')->applyFromArray($styleArray1);
$objPHPExcel->getActiveSheet()->getStyle('J49:J50')->applyFromArray($styleArray1);
$objPHPExcel->getActiveSheet()->getStyle('T47:T48')->applyFromArray($styleArray2);

$objPHPExcel->getActiveSheet()->getStyle('C47:C50')->applyFromArray($styleArray2);
$objPHPExcel->getActiveSheet()->getStyle('J49:J50')->applyFromArray($styleArray2);


												
							



							
                  
$xAxisTickValues9 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$V$54:$AA54', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues9 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$V$55:$AA55', NULL, 5));

//	Build the dataseries


$series9= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues9)-1),			// plotOrder
	$dataseriesLabels9,							// plotLabel
	$xAxisTickValues9,								// plotCategory
	$dataSeriesValues9								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series9->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

$layout9 = new PHPExcel_Chart_Layout();


//	Set the series in the plot area
$plotarea9 = new PHPExcel_Chart_PlotArea($layout9, array($series9));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend9=NULL;
$title9=new PHPExcel_Chart_Title('');   
$yAxisLabel9=NULL;


$objPHPExcel->getActiveSheet()->mergeCells('W53:W53');
$objPHPExcel->getActiveSheet()->mergeCells('W53:W53');
$objPHPExcel->getActiveSheet()->setCellValue('W51', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('Y51', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('W52', '=IF(Z55>V55,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('Y52', '=IF(Z55>Y55,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('W53', '=round((Z55-V55)/ABS(V55),2)');
$objPHPExcel->getActiveSheet()->setCellValue('Y53', '=round((Z55-Y55)/ABS(Y55),2)');
$objPHPExcel->getActiveSheet()->getStyle('Y52')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('W52')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

//conditions
$objConditional1 = new PHPExcel_Style_Conditional();
$objConditional1->setConditionType(PHPExcel_Style_Conditional::CONDITION_CELLIS)
    ->setOperatorType(PHPExcel_Style_Conditional::OPERATOR_LESSTHAN)
    ->addCondition('0');
$objConditional1->getStyle()->getFill()->applyFromArray(
            array(
            'type'       => PHPExcel_Style_Fill::FILL_SOLID,
            'startcolor' => array('argb' => 'FF0000'),
			'endcolor' => array('argb' => 'FF0000')
	));
$objConditional1->getStyle()->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_WHITE);
$objConditional1->getStyle()->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE);


$objConditional3 = new PHPExcel_Style_Conditional();
$objConditional3->setConditionType(PHPExcel_Style_Conditional::CONDITION_CELLIS)
    ->setOperatorType(PHPExcel_Style_Conditional::OPERATOR_GREATERTHANOREQUAL)
    ->addCondition('0');
$objConditional3->getStyle()->getFill()->applyFromArray(
            array(
            'type'       => PHPExcel_Style_Fill::FILL_SOLID,
            'startcolor' => array('argb'=>'008000'),
            'endcolor' => array('argb'=>'008000')
	 ));
	 
$objConditional3->getStyle()->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_WHITE);
$objConditional3->getStyle()->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE);

$objPHPExcel->getActiveSheet()->getStyle('W53')->setConditionalStyles($conditionalStyles);
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('W53')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);

$objPHPExcel->getActiveSheet()->getStyle('Y53')->setConditionalStyles($conditionalStyles);
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('Y53')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);


$objPHPExcel->getActiveSheet()->getStyle('W53')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('Y53')->setConditionalStyles($conditionalStyles);


//conditions




//	Create the chart
$chart9= new PHPExcel_Chart(
	'chart9',		// name
	$title9,		// title
	$legend9,		// legend
	$plotarea9,	// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel9	// yAxisLabel
);

//	Set the position where the chart should appear in the worksheet

$chart9->setTopLeftPosition('V56');
$chart9->setBottomRightPosition('AA62');

//	Add the chart to the worksheet
$objWorksheet->addChart($chart9);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////










// Save Excel 2007 file

echo date('H:i:s') , " Write to Excel2007 format" , EOL;
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->setIncludeCharts(TRUE);
$objWriter->save(str_replace('.php', '.xlsx', __FILE__));
echo date('H:i:s') , " File written to " , str_replace('.php', '.xlsx', pathinfo(__FILE__, PATHINFO_BASENAME)) , EOL;


// Echo memory peak usage
echo date('H:i:s') , " Peak memory usage: " , (memory_get_peak_usage(true) / 1024 / 1024) , " MB" , EOL;

// Echo done
echo date('H:i:s') , " Done writing file" , EOL;
echo 'File has been created in ' , getcwd() , EOL;




?>