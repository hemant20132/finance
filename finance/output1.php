<?php
session_start();
ob_start(); 
$companyname=$_SESSION['companyname'];	
$inputfile=$_SESSION['inputfilename'];
if ($companyname=='' or $inputfile=='')
{
echo "company name or input file not found";
}


error_reporting(E_ALL);
ini_set("display_errors", 1);
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
	
$query1 = "select finyear, totalrevenue_sales, costrevenue,
grossprofit, operatingexp, operatinginc, 
incomebeforetax, netprofit 
from income_statement where company='".$companyname."'";
$result = mysql_query($query1,$conn);

//$result = mysqli_query($conn,$query1);

$i=0;
$j=0;
$results1 = array();
$results2 = array();
$results3 = array();
$results4 = array();
$results5 = array();
$results6 = array();
$results7 = array();

$results27= array();
$results28= array();
$results29= array();
$results30= array();

//while ($info = mysqli_fetch_array($result,MYSQLI_ASSOC))

while ($info = mysql_fetch_array($result,MYSQL_ASSOC))
{
	$finyear=$info['finyear'];
	$revsales=$info['totalrevenue_sales'];
	$costrevenue=$info['costrevenue'];
	$grossprofit=$info['grossprofit'];
	$operatingexp=$info['operatingexp'];
	$operatinginc=$info['operatinginc'];
	$incomebeforetax=$info['incomebeforetax'];
	$netprofit=$info['netprofit'];
	
	$grossprofitpercentage=round($info['grossprofit']/$info['totalrevenue_sales']*100,2);
	$operatingincpercentage=round($info['operatinginc']/$info['totalrevenue_sales']*100,2);
	$incomebeforetaxpercentage=round($info['incomebeforetax']/$info['totalrevenue_sales']*100,2);	
	$netprofitpercentage=round($info['netprofit']/$info['totalrevenue_sales']*100,2);	
	
	$results1[$i][$j]=$finyear;
	$results2[$i][$j]=$finyear;
	$results3[$i][$j]=$finyear;
	$results4[$i][$j]=$finyear;
	$results5[$i][$j]=$finyear;
	$results6[$i][$j]=$finyear;
	$results7[$i][$j]=$finyear;

	$results27[$i][$j]=$finyear;
	$results28[$i][$j]=$finyear;
	$results29[$i][$j]=$finyear;
	$results30[$i][$j]=$finyear;
			
	$j=$j+1;	
	
	$results1[$i+1][$j]=$revsales;
	$results2[$i+1][$j]=$costrevenue;
	$results3[$i+1][$j]=$grossprofit;
	$results4[$i+1][$j]=$operatingexp;
	$results5[$i+1][$j]=$operatinginc;
	$results6[$i+1][$j]=$incomebeforetax;
	$results7[$i+1][$j]=$netprofit;

	$results27[$i+1][$j]=$grossprofitpercentage;
	$results28[$i+1][$j]=$operatingincpercentage;
	$results29[$i+1][$j]=$incomebeforetaxpercentage;
	$results30[$i+1][$j]=$netprofitpercentage;
}


$query3 = "select finyear, cashandequivalents, accountrecive, inventory,
totalcurrenassets, propertyequip, longterminvst,totallongterm, totalassets,
accountpayable, totalcurliab, longtermdebt, longtermliab, totalliabilities,
totalequity, liabilitiesandequity  
from balance_sheet where company='".$companyname."'";

$resultbalance = mysql_query($query3,$conn);

//$resultbalance = mysqli_query($conn,$query3);
$results8 = array();
$results9 = array();
$results10 = array();
$results11 = array();
$results12 = array();
$results13 = array();
$results14 = array();
$results15 = array();
$results16 = array();
$results17 = array();
$results18 = array();
$results19 = array();
$results20 = array();
$results21 = array();
$results22 = array();

$results31 = array();
$results32 = array();
$results38 = array();
$results39 = array();

$results41= array();




//while ($info1 = mysqli_fetch_array($resultbalance,MYSQLI_ASSOC))

while ($info1 = mysql_fetch_array($resultbalance,MYSQL_ASSOC))
{
	$finyear = $info1['finyear'];
	$cashandequivalents = $info1['cashandequivalents'];
	$accountreceive = $info1['accountrecive'];
	$inventory = $info1['inventory'];
	$totalcurrentassets= $info1['totalcurrenassets'];
	$propertyequip= $info1['propertyequip'];
	$longterminvst= $info1['longterminvst'];
	$totallongterm= $info1['totallongterm'];
	$totalassets= $info1['totalassets'];
	$accountpayable= $info1['accountpayable'];
	$totalcurliab = $info1['totalcurliab'];
	$longtermdebt = $info1['longtermdebt'];
	$totalequity = $info1['totalequity'];
	$totalliabilities= $info1['totalliabilities'];
	$totalequity=$info1['totalequity'];
	$liabilitiesandequity=$info1['liabilitiesandequity'];  
	$currentratio=$info1['totalcurrenassets']/$info1['totalcurliab'];
	$quickratio=($info1['totalcurrenassets']-$info1['inventory'])/$info1['totalcurliab'];
	$debtratio=($info1['totalliabilities']/$info1['totalassets']);
	$debtequityratio=($info1['totalliabilities']/$info1['totalequity']);
	
	
	$results8[$i][$j]=$finyear;
	$results9[$i][$j]=$finyear;
	$results10[$i][$j]=$finyear;
	$results11[$i][$j]=$finyear;
	$results12[$i][$j]=$finyear;
	$results13[$i][$j]=$finyear;
	$results14[$i][$j]=$finyear;
	$results15[$i][$j]=$finyear;
	$results16[$i][$j]=$finyear;
	$results17[$i][$j]=$finyear;
	$results18[$i][$j]=$finyear;
	$results19[$i][$j]=$finyear;
	$results20[$i][$j]=$finyear;
	$results21[$i][$j]=$finyear;
	$results22[$i][$j]=$finyear;
	
	$results31[$i][$j]=$finyear;
	$results32[$i][$j]=$finyear;
	
	$results38[$i][$j]=$finyear;
	$results39[$i][$j]=$finyear;
	$results41[$i][$j]=$finyear;
	
	$j=$j+1;	
	
	$results8[$i+1][$j]=$cashandequivalents;
	$results9[$i+1][$j]=$accountreceive;
	$results10[$i+1][$j]=$inventory;
	$results11[$i+1][$j]=$totalcurrentassets;
	$results12[$i+1][$j]=$propertyequip;
	$results13[$i+1][$j]= $longterminvst;
	$results14[$i+1][$j]= $totallongterm;
	$results15[$i+1][$j]= $totalassets;
	$results16[$i+1][$j]= $accountpayable;
	$results17[$i+1][$j]= $totalcurliab;
	$results18[$i+1][$j]= $longtermdebt;
	$results19[$i+1][$j]= $totalequity;
	$results20[$i+1][$j]= $totalliabilities;
	$results21[$i+1][$j]= $totalequity;
	$results22[$i+1][$j]= $liabilitiesandequity;

	$results31[$i+1][$j]= $currentratio;
	$results32[$i+1][$j]= $quickratio;
	$results38[$i+1][$j]= $debtratio;
	$results39[$i+1][$j]= $debtequityratio;
	$results41[$i+1][$j]= ($totalcurrentassets-$totalcurliab)/$totalassets;
}

$query4 = "select finyear, netchangeincash, cashfromoper, cashfrominvesting, cashfromfinance
from cash_flow where company='". $companyname ."'";

$resultcash = mysql_query($query4, $conn);
//$resultcash = mysqli_query($conn,$query4);

$results23 = array();
$results24 = array();
$results25 = array();
$results26 = array();

//while ($info4 = mysqli_fetch_array($resultcash,MYSQLI_ASSOC))

while ($info4 = mysql_fetch_array($resultcash,MYSQL_ASSOC))
{
	$finyear = $info4['finyear'];
	$netchangeincash=$info4['netchangeincash'];
	$cashfromoper = $info4['cashfromoper'];
	$cashfrominvesting = $info4['cashfrominvesting'];
	$cashfromfinance = $info4['cashfromfinance'];
	
	$results23[$i][$j]=$finyear;
	$results24[$i][$j]=$finyear;
	$results25[$i][$j]=$finyear;
	$results26[$i][$j]=$finyear;	
	
	$j=$j+1;	
	
	$results23[$i+1][$j]=$cashfromoper;
	$results24[$i+1][$j]=$cashfrominvesting;
	$results25[$i+1][$j]=$cashfromfinance;
	$results26[$i+1][$j]=$netchangeincash;
}


$query5 = "select income_statement.finyear, income_statement.totalrevenue_sales,income_statement.costrevenue,
income_statement.netprofit, 
balance_sheet.totalassets, balance_sheet.inventory, balance_sheet.accountrecive, balance_sheet.totalequity from income_statement, 
balance_sheet where income_statement.company='" . $companyname . "' and income_statement.finyear=balance_sheet.finyear ";


$resultnew = mysql_query($query5, $conn);
//$resultnew = mysqli_query($conn,$query5);
$results33 = array();
$results34 = array();
$results35 = array();
$results36 = array();
$results37 = array();

$results40 = array();
$results42 = array();


//while ($info5 = mysqli_fetch_array($resultnew,MYSQLI_ASSOC))
$i=0;
$j=0;
while ($info5 = mysql_fetch_array($resultnew,MYSQL_ASSOC))
{
	$finyear = $info5['finyear'];
	$revsales = $info5['totalrevenue_sales'];
	$costrevenue =$info5['costrevenue'];
	$netprofit= $info5['netprofit'];
	$totalassets =$info5['totalassets'];
	$inventory =$info5['inventory'];
	$accountreceive=$info5['accountrecive'];
	$totalequity=$info5['totalequity'];
	
	$results33[$i][$j] = $finyear;
	$results34[$i][$j] = $finyear;
	$results35[$i][$j] = $finyear;
	$results36[$i][$j] = $finyear;
	$results37[$i][$j] = $finyear;
	$results40[$i][$j] = $finyear;
	$results42[$i][$j] = $finyear;
	
	$j=$j+1;	

	$results33[$i+1][$j]=$revsales/$totalassets;
	$results34[$i+1][$j]=$costrevenue/$inventory;
	$results35[$i+1][$j]=$revsales/$accountreceive;
	$results36[$i+1][$j]=round($accountreceive/($revsales/365),0);
	$results37[$i+1][$j]=round($inventory/($costrevenue/365),0);
	$results40[$i+1][$j]=$netprofit/$totalassets;
	$results42[$i+1][$j]=round($netprofit/$totalequity*100,2);
	}
 

$objPHPExcel = new PHPExcel();

$newSheet=new PHPExcel_Worksheet($objPHPExcel,'Financial_Ratios');
$objPHPExcel->addSheet($newSheet,0);
$objWorksheet=$objPHPExcel->setActiveSheetIndex(0);  
$objPHPExcel->getActiveSheet()->getDefaultColumnDimension()->setWidth(9);

$styleArray = array(
	'borders' => array(
		'outline' => array(
			'style' => PHPExcel_Style_Border::BORDER_THIN,
			'color' => array('argb' => 'ffffff'),
		),
	),
);

$objWorksheet->getStyle('A1:T480')->applyFromArray($styleArray);

$objWorksheet->fromArray(
$results1,NULL,'C8'
);

///// financial ratios starts here/////////
$objPHPExcel->getActiveSheet()->mergeCells('A1:T1');
$objPHPExcel->getActiveSheet()->setCellValue('A1', 'Financial Ratios');
$objPHPExcel->getActiveSheet()->getStyle("A1")->getFont()->setBold(true)
                                ->setName('Verdana')
                                ->setSize(20)
								->getColor()->setRGB('FFFFF');

$objPHPExcel->getActiveSheet()->getStyle('A1:T1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$sheet = $objPHPExcel->getActiveSheet();

$sheet->getStyle('A1:T1')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->getStyle('A1:T1')->getFill()->getStartColor()->setRGB('6a5acd');

$dataseriesLabels1 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$objPHPExcel->getActiveSheet()->setCellValue('C7', 'Sales');
$objPHPExcel->getActiveSheet()->getStyle("C7")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');
                  
$xAxisTickValues1 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$8:$G8', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues1 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$9:$G9', NULL, 5));

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

//	Set the series in the plot area
$plotarea1 = new PHPExcel_Chart_PlotArea(NULL, array($series1));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend1=NULL;
$title1 = new PHPExcel_Chart_Title('Sales');
$yAxisLabel1= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P9:P10');
$objPHPExcel->getActiveSheet()->mergeCells('R9:R10');
$objPHPExcel->getActiveSheet()->setCellValue('P7', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R7', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P8', '=IF(G8>C8,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P9', '=round((G9-C9)/ABS(C9),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R9', '=round((G9-F9)/ABS(C9),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R8', '=IF(G8>C8,"↑","↓")');
$objPHPExcel->getActiveSheet()->getStyle('P8')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R8')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P9')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P9')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R9')->setConditionalStyles($conditionalStyles);


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

//	Set the position where the chart should appear in the Financial_Ratios

$chart1->setTopLeftPosition('I4');
$chart1->setBottomRightPosition('O15');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart1);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////

$objWorksheet->fromArray(
$results3,NULL,'C22'
);

$dataseriesLabels2 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$objPHPExcel->getActiveSheet()->setCellValue('C21', 'Gross Profit');

$objPHPExcel->getActiveSheet()->getStyle("C21")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues2 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$22:$G22', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues2 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$23:$G23', NULL, 5));

//	Build the dataseries
$series2= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues2)-1),			// plotOrder
	$dataseriesLabels2,								// plotLabel
	$xAxisTickValues2,								// plotCategory
	$dataSeriesValues2								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series2->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea2 = new PHPExcel_Chart_PlotArea(NULL, array($series2));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend2=NULL;
$title2 = new PHPExcel_Chart_Title('Gross Profit');
$yAxisLabel2= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P23:P24');
$objPHPExcel->getActiveSheet()->mergeCells('R23:R24');
$objPHPExcel->getActiveSheet()->setCellValue('P21', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R21', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P22', '=IF(G23>C23,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P23', '=round((G23-C23)/ABS(C23),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R23', '=round((G23-F23)/ABS(C23),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R22', '=IF(G23>C23,"↑","↓")');
$objPHPExcel->getActiveSheet()->getStyle('P22')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R22')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P23')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P23')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R23')->setConditionalStyles($conditionalStyles);


//conditions



//	Create the chart
$chart2= new PHPExcel_Chart(
	'chart2',		// name
	$title2,		// title
	$legend2,		// legend
	$plotarea2,		// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel2	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart2->setTopLeftPosition('I18');
$chart2->setBottomRightPosition('O28');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart2);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////


$objWorksheet->fromArray(
$results5,NULL,'C46'
);

$dataseriesLabels3 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$objPHPExcel->getActiveSheet()->setCellValue('C45', 'Operating Income EBIT');
$objPHPExcel->getActiveSheet()->getStyle("C45")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');


$xAxisTickValues3 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$46:$G46', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues3 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$47:$G47', NULL, 5));

//	Build the dataseries
$series3= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues3)-1),			// plotOrder
	$dataseriesLabels3,								// plotLabel
	$xAxisTickValues3,								// plotCategory
	$dataSeriesValues3								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series3->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea3 = new PHPExcel_Chart_PlotArea(NULL, array($series3));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend3=NULL;
$title3 = new PHPExcel_Chart_Title('Operating Income (EBIT)');
$yAxisLabel3= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P47:P48');
$objPHPExcel->getActiveSheet()->mergeCells('R47:R48');
$objPHPExcel->getActiveSheet()->setCellValue('P45', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R45', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P46', '=IF(G47>C47,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P47', '=round((G47-C47)/ABS(C47),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R47', '=round((G47-F47)/ABS(F47),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R46', '=IF(G47>C47,"↑","↓")');
$objPHPExcel->getActiveSheet()->getStyle('P46')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R46')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P47')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P47')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R47')->setConditionalStyles($conditionalStyles);
//conditions

//	Create the chart
$chart3= new PHPExcel_Chart(
	'chart3',		// name
	$title3,		// title
	$legend3,		// legend
	$plotarea3,		// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel3	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart3->setTopLeftPosition('I41');
$chart3->setBottomRightPosition('O51');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart3);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////



$objWorksheet->fromArray(
$results6,NULL,'C67'
);

$dataseriesLabels4 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$objPHPExcel->getActiveSheet()->setCellValue('C66', 'Income Before Tax');

$objPHPExcel->getActiveSheet()->getStyle("C66")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');
$xAxisTickValues4 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$67:$G67', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues4 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$68:$G68', NULL, 5));

//	Build the dataseries
$series4= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues4)-1),			// plotOrder
	$dataseriesLabels4,								// plotLabel
	$xAxisTickValues4,								// plotCategory
	$dataSeriesValues4								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series4->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea4 = new PHPExcel_Chart_PlotArea(NULL, array($series4));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend4=NULL;
$title4 = new PHPExcel_Chart_Title('Income Before Tax');

$yAxisLabel4= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P68:P69');
$objPHPExcel->getActiveSheet()->mergeCells('R68:R69');
$objPHPExcel->getActiveSheet()->setCellValue('P66', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R66', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P67', '=IF(G68>C68,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R67', '=IF(G68>F68,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P68', '=round((G68-C68)/ABS(C68),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R68', '=round((G68-F68)/ABS(F68),2)');
$objPHPExcel->getActiveSheet()->getStyle('P67')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R67')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P68')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P68')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R68')->setConditionalStyles($conditionalStyles);
//conditions





//	Create the chart
$chart4= new PHPExcel_Chart(
	'chart4',		// name
	$title4,		// title
	$legend4,		// legend
	$plotarea4,		// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel4	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart4->setTopLeftPosition('I63');
$chart4->setBottomRightPosition('O73');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart4);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////


$objWorksheet->fromArray(
$results7,NULL,'C87'
);

$dataseriesLabels5 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$objPHPExcel->getActiveSheet()->setCellValue('C86', 'Net Profit');
$objPHPExcel->getActiveSheet()->getStyle("C86")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues5 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$67:$G67', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues5 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$87:$G87', NULL, 5));

//	Build the dataseries
$series5= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues5)-1),			// plotOrder
	$dataseriesLabels5,								// plotLabel
	$xAxisTickValues5,								// plotCategory
	$dataSeriesValues5								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series5->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea5 = new PHPExcel_Chart_PlotArea(NULL, array($series5));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend5=NULL;
$title5 = new PHPExcel_Chart_Title('Net Profit');
$yAxisLabel5= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P88:P89');
$objPHPExcel->getActiveSheet()->mergeCells('R88:R89');
$objPHPExcel->getActiveSheet()->setCellValue('P86', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R86', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P87', '=IF(G88>C88,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R87', '=IF(G88>F88,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P88', '=round((G88-C88)/ABS(C88),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R88', '=round((G88-F88)/ABS(F88),2)');
$objPHPExcel->getActiveSheet()->getStyle('P87')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R87')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P88')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P88')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R88')->setConditionalStyles($conditionalStyles);
//conditions




//	Create the chart
$chart5= new PHPExcel_Chart(
	'chart5',		// name
	$title5,		// title
	$legend5,		// legend
	$plotarea5,		// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel5	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart5->setTopLeftPosition('I83');
$chart5->setBottomRightPosition('O93');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart5);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////




$objWorksheet->fromArray(
$results8,NULL,'C107'
);

$dataseriesLabels6 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$objPHPExcel->getActiveSheet()->setCellValue('C106', 'Cash &  Equivalents');
$objPHPExcel->getActiveSheet()->getStyle("C106")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');


$xAxisTickValues6 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$107:$G107', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues6 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$108:$G108', NULL, 5));

//	Build the dataseries
$series6= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues5)-1),			// plotOrder
	$dataseriesLabels6,								// plotLabel
	$xAxisTickValues6,								// plotCategory
	$dataSeriesValues6								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series6->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea6 = new PHPExcel_Chart_PlotArea(NULL, array($series6));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend6=NULL;
$title6 = new PHPExcel_Chart_Title('Cash &  Equivalents');
$yAxisLabel6= NULL;


$objPHPExcel->getActiveSheet()->mergeCells('P108:P109');
$objPHPExcel->getActiveSheet()->mergeCells('R108:R109');
$objPHPExcel->getActiveSheet()->setCellValue('P106', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R106', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P107', '=IF(G108>C108,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R107', '=IF(G108>F108,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P108', '=round((G108-C108)/ABS(C108),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R108', '=round((G108-F108)/ABS(F108),2)');
$objPHPExcel->getActiveSheet()->getStyle('P107')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R107')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P108')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P108')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R108')->setConditionalStyles($conditionalStyles);
//conditions


//	Create the chart
$chart6= new PHPExcel_Chart(
	'chart6',		// name
	$title6,		// title
	$legend6,		// legend
	$plotarea6,		// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel6	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart6->setTopLeftPosition('I105');
$chart6->setBottomRightPosition('O115');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart6);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////

$objWorksheet->fromArray(
$results9,NULL,'C122'
);

$dataseriesLabels7 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C121', 'Account Receivable');
$objPHPExcel->getActiveSheet()->getStyle("C121")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues7 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$117:$G117', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues7 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$117:$G117', NULL, 5));

//	Build the dataseries
$series7= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues7)-1),			// plotOrder
	$dataseriesLabels7,								// plotLabel
	$xAxisTickValues7,								// plotCategory
	$dataSeriesValues7								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series7->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea7 = new PHPExcel_Chart_PlotArea(NULL, array($series6));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend7=NULL;
$title7 = new PHPExcel_Chart_Title('Account Receivable');
$yAxisLabel7= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P123:P124');
$objPHPExcel->getActiveSheet()->mergeCells('R123:R124');
$objPHPExcel->getActiveSheet()->setCellValue('P121', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R121', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P122', '=IF(G108>C108,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R122', '=IF(G108>F108,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P123', '=round((G108-C108)/ABS(C108),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R123', '=round((G108-F108)/ABS(F108),2)');
$objPHPExcel->getActiveSheet()->getStyle('P122')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R122')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P123')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P123')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R123')->setConditionalStyles($conditionalStyles);
//conditions


//	Create the chart
$chart7= new PHPExcel_Chart(
	'chart7',		// name
	$title7,		// title
	$legend7,		// legend
	$plotarea7,		// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel7	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart7->setTopLeftPosition('I119');
$chart7->setBottomRightPosition('O129');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart7);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////



$objWorksheet->fromArray(
$results10,NULL,'C136'
);

$dataseriesLabels8 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C135', 'Inventory');
$objPHPExcel->getActiveSheet()->getStyle("C135")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues8 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$136:$G136', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues8 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$137:$G137', NULL, 5));

//	Build the dataseries
$series8= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues8)-1),			// plotOrder
	$dataseriesLabels8,								// plotLabel
	$xAxisTickValues8,								// plotCategory
	$dataSeriesValues8								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series8->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea8 = new PHPExcel_Chart_PlotArea(NULL, array($series8));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend8=NULL;
$title8 = new PHPExcel_Chart_Title('Inventory');
$yAxisLabel8= NULL;


$objPHPExcel->getActiveSheet()->mergeCells('P137:P138');
$objPHPExcel->getActiveSheet()->mergeCells('R137:R138');
$objPHPExcel->getActiveSheet()->setCellValue('P135', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R135', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P136', '=IF(G108>C108,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R136', '=IF(G108>F108,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P137', '=round((G137-C137)/ABS(C137),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R137', '=round((G137-F137)/ABS(F137),2)');
$objPHPExcel->getActiveSheet()->getStyle('P136')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R136')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P137')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P137')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R137')->setConditionalStyles($conditionalStyles);
//conditions


//	Create the chart
$chart8= new PHPExcel_Chart(
	'chart8',		// name
	$title8,		// title
	$legend8,		// legend
	$plotarea8,		// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel8	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart8->setTopLeftPosition('I131');
$chart8->setBottomRightPosition('O141');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart8);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////


$objWorksheet->fromArray(
$results13,NULL,'C150'
);

$dataseriesLabels9 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C149', 'Long Term Investments');
$objPHPExcel->getActiveSheet()->getStyle("C149")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues9 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$150:$G150', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues9 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$151:$G151', NULL, 5));

//	Build the dataseries
$series9= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues9)-1),			// plotOrder
	$dataseriesLabels9,								// plotLabel
	$xAxisTickValues9,								// plotCategory
	$dataSeriesValues9								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series9->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea9 = new PHPExcel_Chart_PlotArea(NULL, array($series9));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend9=NULL;
$title9 = new PHPExcel_Chart_Title('Long Term Investments');
$yAxisLabel9= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P151:P152');
$objPHPExcel->getActiveSheet()->mergeCells('R151:R152');
$objPHPExcel->getActiveSheet()->setCellValue('P149', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R149', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P150', '=IF(G151>C151,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R150', '=IF(G151>F151,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P151', '=round((G151-C151)/ABS(C151),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R151', '=round((G151-F151)/ABS(F151),2)');
$objPHPExcel->getActiveSheet()->getStyle('P150')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R150')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P151')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P151')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R151')->setConditionalStyles($conditionalStyles);
//conditions


//	Create the chart
$chart9= new PHPExcel_Chart(
	'chart9',		// name
	$title9,		// title
	$legend9,		// legend
	$plotarea9,		// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel9	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart9->setTopLeftPosition('I145');
$chart9->setBottomRightPosition('O155');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart9);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////



$objWorksheet->fromArray(
$results15,NULL,'C163'
);

$dataseriesLabels10 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C162', 'Total Assets');
$objPHPExcel->getActiveSheet()->getStyle("C162")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues10 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$163:$G163', NULL, 5),//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues10 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$164:$G164', NULL, 5));

//	Build the dataseries
$series10= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues10)-1),			// plotOrder
	$dataseriesLabels10,								// plotLabel
	$xAxisTickValues10,								// plotCategory
	$dataSeriesValues10								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series10->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea10 = new PHPExcel_Chart_PlotArea(NULL, array($series10));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend10=NULL;
$title10 = new PHPExcel_Chart_Title('Total Assets');
$yAxisLabel10= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P164:P165');
$objPHPExcel->getActiveSheet()->mergeCells('R164:R165');
$objPHPExcel->getActiveSheet()->setCellValue('P162', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R162', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P163', '=IF(G164>C164,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R163', '=IF(G164>F164,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P164', '=round((G164-C164)/ABS(C164),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R164', '=round((G164-F164)/ABS(F164),2)');
$objPHPExcel->getActiveSheet()->getStyle('P163')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R163')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P164')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P164')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R164')->setConditionalStyles($conditionalStyles);
//conditions


//	Create the chart
$chart10= new PHPExcel_Chart(
	'chart10',		// name
	$title10,		// title
	$legend10,		// legend
	$plotarea10,		// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel10	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart10->setTopLeftPosition('I159');
$chart10->setBottomRightPosition('O169');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart10);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////


$objWorksheet->fromArray(
$results16,NULL,'C176'
);

$dataseriesLabels11 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C175', 'Accounts Payable');
$objPHPExcel->getActiveSheet()->getStyle("C175")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues11 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$176:$G176', NULL, 5),//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues11 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$177:$G177', NULL, 5));

//	Build the dataseries
$series11= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues11)-1),			// plotOrder
	$dataseriesLabels11,								// plotLabel
	$xAxisTickValues11,								// plotCategory
	$dataSeriesValues11								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series11->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea11 = new PHPExcel_Chart_PlotArea(NULL, array($series11));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend11=NULL;
$title11 = new PHPExcel_Chart_Title('Accounts Payable');
$yAxisLabel11= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P177:P178');
$objPHPExcel->getActiveSheet()->mergeCells('R177:R178');
$objPHPExcel->getActiveSheet()->setCellValue('P175', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R175', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P176', '=IF(G177>C177,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R176', '=IF(G177>F177,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P177', '=round((G177-C177)/ABS(C177),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R177', '=round((G177-F177)/ABS(F177),2)');
$objPHPExcel->getActiveSheet()->getStyle('P176')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R176')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P177')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P177')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R177')->setConditionalStyles($conditionalStyles);
//conditions


//	Create the chart
$chart11= new PHPExcel_Chart(
	'chart11',		// name
	$title11,		// title
	$legend11,		// legend
	$plotarea11,		// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel11	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart11->setTopLeftPosition('I173');
$chart11->setBottomRightPosition('O183');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart11);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////


$objWorksheet->fromArray(
$results17,NULL,'C191'
);

$dataseriesLabels12 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C190', 'Current Liabilities');
$objPHPExcel->getActiveSheet()->getStyle("C190")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues12 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$191:$G191', NULL, 5),//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues12 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$192:$G192', NULL, 5));

//	Build the dataseries
$series12= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues12)-1),			// plotOrder
	$dataseriesLabels12,								// plotLabel
	$xAxisTickValues12,								// plotCategory
	$dataSeriesValues12								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series12->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea12 = new PHPExcel_Chart_PlotArea(NULL, array($series12));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend12=NULL;
$title12 = new PHPExcel_Chart_Title('Current Liabilities');
$yAxisLabel12= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P192:P193');
$objPHPExcel->getActiveSheet()->mergeCells('R192:R193');
$objPHPExcel->getActiveSheet()->setCellValue('P190', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R190', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P191', '=IF(G192>C191,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R191', '=IF(G192>F191,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P192', '=round((G192-C192)/ABS(C192),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R192', '=round((G192-F192)/ABS(F192),2)');
$objPHPExcel->getActiveSheet()->getStyle('P191')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R191')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P192')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P192')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R192')->setConditionalStyles($conditionalStyles);
//conditions


//	Create the chart
$chart12= new PHPExcel_Chart(
	'chart12',		// name
	$title12,		// title
	$legend12,		// legend
	$plotarea12,		// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel12	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart12->setTopLeftPosition('I187');
$chart12->setBottomRightPosition('O197');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart12);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////



$objWorksheet->fromArray(
$results18, NULL, 'C206'
);

$dataseriesLabels13 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C205', 'Long Term Debt');
$objPHPExcel->getActiveSheet()->getStyle("C205")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues13 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$206:$G206', NULL, 5),//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues13 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$207:$G207', NULL, 5));

//	Build the dataseries
$series13= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues13)-1),			// plotOrder
	$dataseriesLabels13,								// plotLabel
	$xAxisTickValues13,								// plotCategory
	$dataSeriesValues13								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series13->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea13 = new PHPExcel_Chart_PlotArea(NULL, array($series13));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend13=NULL;
$title13 = new PHPExcel_Chart_Title('Long Term Debt');
$yAxisLabel13= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P207:P208');
$objPHPExcel->getActiveSheet()->mergeCells('R207:R208');
$objPHPExcel->getActiveSheet()->setCellValue('P205', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R205', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P206', '=IF(G207>C207,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R206', '=IF(G207>F207,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P207', '=round((G207-C207)/ABS(C207),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R207', '=round((G207-F207)/ABS(F207),2)');
$objPHPExcel->getActiveSheet()->getStyle('P206')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R206')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P207')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P207')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R207')->setConditionalStyles($conditionalStyles);
//conditions


//	Create the chart
$chart13= new PHPExcel_Chart(
	'chart13',		// name
	$title13,		// title
	$legend13,		// legend
	$plotarea13,		// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel13	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart13->setTopLeftPosition('I201');
$chart13->setBottomRightPosition('O210');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart13);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////



$objWorksheet->fromArray(
$results20, NULL, 'C219'
);

$dataseriesLabels14 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C218', 'Total Liabilities');
$objPHPExcel->getActiveSheet()->getStyle("C218")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues14 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$219:$G219', NULL, 5),//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues14 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$220:$G220', NULL, 5));

//	Build the dataseries
$series14= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues14)-1),			// plotOrder
	$dataseriesLabels14,								// plotLabel
	$xAxisTickValues14,								// plotCategory
	$dataSeriesValues14								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series14->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea14 = new PHPExcel_Chart_PlotArea(NULL, array($series14));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend14=NULL;
$title14 = new PHPExcel_Chart_Title('Total Liabilities');
$yAxisLabel14= NULL;


$objPHPExcel->getActiveSheet()->mergeCells('P220:P221');
$objPHPExcel->getActiveSheet()->mergeCells('R220:R221');
$objPHPExcel->getActiveSheet()->setCellValue('P218', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R218', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P219', '=IF(G220>C220,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R219', '=IF(G220>F220,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P220', '=round((G220-C220)/ABS(C220),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R220', '=round((G220-F220)/ABS(F220),2)');
$objPHPExcel->getActiveSheet()->getStyle('P219')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R219')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P220')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P220')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R220')->setConditionalStyles($conditionalStyles);
//conditions



//	Create the chart
$chart14= new PHPExcel_Chart(
	'chart14',		// name
	$title14,		// title
	$legend14,		// legend
	$plotarea14,		// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel14	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart14->setTopLeftPosition('I214');
$chart14->setBottomRightPosition('O224');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart14);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////


$objWorksheet->fromArray(
$results21, NULL, 'C232'
);

$dataseriesLabels15 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C231', 'Total Equity');
$objPHPExcel->getActiveSheet()->getStyle("C231")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues15 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$232:$G232', NULL, 5),//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues15 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$233:$G233', NULL, 5));

//	Build the dataseries
$series15= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues15)-1),			// plotOrder
	$dataseriesLabels15,								// plotLabel
	$xAxisTickValues15,								// plotCategory
	$dataSeriesValues15								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series15->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea15 = new PHPExcel_Chart_PlotArea(NULL, array($series14));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend15=NULL;
$title15 = new PHPExcel_Chart_Title('Total Equity');
$yAxisLabel15= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P233:P234');
$objPHPExcel->getActiveSheet()->mergeCells('R233:R234');
$objPHPExcel->getActiveSheet()->setCellValue('P231', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R231', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P232', '=IF(G233>C233,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R232', '=IF(G233>F233,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P233', '=round((G233-C233)/ABS(C233),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R233', '=round((G233-F233)/ABS(F233),2)');
$objPHPExcel->getActiveSheet()->getStyle('P232')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R232')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P233')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P233')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R233')->setConditionalStyles($conditionalStyles);
//conditions


//	Create the chart
$chart15= new PHPExcel_Chart(
	'chart15',		// name
	$title15,		// title
	$legend15,		// legend
	$plotarea15,		// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel15	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart15->setTopLeftPosition('I228');
$chart15->setBottomRightPosition('O238');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart15);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////


$objWorksheet->fromArray(
$results22, NULL, 'C247'
);

$dataseriesLabels16 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C246', 'Total Liabilities & Equity');
$objPHPExcel->getActiveSheet()->getStyle("C246")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues16 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$247:$G247', NULL, 5),//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues16 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$248:$G248', NULL, 5));

//	Build the dataseries
$series16= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues16)-1),			// plotOrder
	$dataseriesLabels16,								// plotLabel
	$xAxisTickValues16,								// plotCategory
	$dataSeriesValues16								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series16->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea16 = new PHPExcel_Chart_PlotArea(NULL, array($series16));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend16=NULL;
$title16 = new PHPExcel_Chart_Title('Total Liabilities & Equity');

$yAxisLabel16= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P248:P249');
$objPHPExcel->getActiveSheet()->mergeCells('R248:R249');
$objPHPExcel->getActiveSheet()->setCellValue('P246', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R246', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P247', '=IF(G248>C248,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R247', '=IF(G248>F248,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P248', '=round((G248-C248)/ABS(C248),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R248', '=round((G248-F248)/ABS(F248),2)');
$objPHPExcel->getActiveSheet()->getStyle('P247')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R247')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P248')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P248')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R248')->setConditionalStyles($conditionalStyles);
//conditions


//	Create the chart
$chart16= new PHPExcel_Chart(
	'chart16',		// name
	$title16,		// title
	$legend16,		// legend
	$plotarea16,		// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel16	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart16->setTopLeftPosition('I242');
$chart16->setBottomRightPosition('O252');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart16);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////


$objWorksheet->fromArray(
$results22, NULL, 'C261'
);

$dataseriesLabels17 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C260', 'Cash from Operating Activities');
$objPHPExcel->getActiveSheet()->getStyle("C260")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues17 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$261:$G261', NULL, 5),//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues17 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$262:$G262', NULL, 5));

//	Build the dataseries
$series17= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues17)-1),			// plotOrder
	$dataseriesLabels17,								// plotLabel
	$xAxisTickValues17,								// plotCategory
	$dataSeriesValues17								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series17->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea17 = new PHPExcel_Chart_PlotArea(NULL, array($series17));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend17=NULL;
$title17 = new PHPExcel_Chart_Title('Cash From Operating Activities');
$yAxisLabel17= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P262:P263');
$objPHPExcel->getActiveSheet()->mergeCells('R262:R263');
$objPHPExcel->getActiveSheet()->setCellValue('P260', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R260', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P261', '=IF(G262>C262,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R261', '=IF(G262>F262,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P262', '=round((G262-C262)/ABS(C262),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R262', '=round((G262-F262)/ABS(F262),2)');
$objPHPExcel->getActiveSheet()->getStyle('P261')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R261')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P262')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P262')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R262')->setConditionalStyles($conditionalStyles);
//conditions

//	Create the chart
$chart17= new PHPExcel_Chart(
	'chart17',		// name
	$title17,		// title
	$legend17,		// legend
	$plotarea17,		// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel17	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart17->setTopLeftPosition('I256');
$chart17->setBottomRightPosition('O266');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart17);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////



$objWorksheet->fromArray(
$results22, NULL, 'C261'
);

$dataseriesLabels18 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C260', 'Net Change In Cash');
$objPHPExcel->getActiveSheet()->getStyle("C260")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues18 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$261:$G261', NULL, 5),//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues18 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$262:$G262', NULL, 5));

//	Build the dataseries
$series18= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues18)-1),			// plotOrder
	$dataseriesLabels18,								// plotLabel
	$xAxisTickValues18,								// plotCategory
	$dataSeriesValues18								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series18->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea18 = new PHPExcel_Chart_PlotArea(NULL, array($series18));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend18=NULL;
$title18 = new PHPExcel_Chart_Title('Net Change in Cash');
$yAxisLabel18= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P276:P277');
$objPHPExcel->getActiveSheet()->mergeCells('R276:R277');
$objPHPExcel->getActiveSheet()->setCellValue('P274', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R274', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P275', '=IF(G276>C276,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R275', '=IF(G276>F276,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P276', '=round((G276-C276)/ABS(C276),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R276', '=round((G276-F276)/ABS(F276),2)');
$objPHPExcel->getActiveSheet()->getStyle('P275')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R275')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P276')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P276')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R276')->setConditionalStyles($conditionalStyles);
//conditions

//	Create the chart
$chart18= new PHPExcel_Chart(
	'chart18',		// name
	$title18,		// title
	$legend18,		// legend
	$plotarea18,		// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel18	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart18->setTopLeftPosition('I256');
$chart18->setBottomRightPosition('O266');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart18);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////



$objWorksheet->fromArray(
$results23, NULL, 'C261'
);

$dataseriesLabels19 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C260', 'Cash from Operating Activities');
$objPHPExcel->getActiveSheet()->getStyle("C260")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues19 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$261:$G261', NULL, 5),//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues19 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$262:$G262', NULL, 5));

//	Build the dataseries
$series19= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues19)-1),			// plotOrder
	$dataseriesLabels19,								// plotLabel
	$xAxisTickValues19,								// plotCategory
	$dataSeriesValues19								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series19->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea19 = new PHPExcel_Chart_PlotArea(NULL, array($series19));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend19=NULL;
$title19 = new PHPExcel_Chart_Title('Cash from Operating Activities');
$yAxisLabel19= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P290:P291');
$objPHPExcel->getActiveSheet()->mergeCells('R290:R291');
$objPHPExcel->getActiveSheet()->setCellValue('P288', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R288', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P289', '=IF(G290>C290,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R289', '=IF(G290>F290,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P290', '=round((G290-C290)/ABS(C290),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R290', '=round((G290-F290)/ABS(F290),2)');
$objPHPExcel->getActiveSheet()->getStyle('P289')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R289')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P290')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P290')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R290')->setConditionalStyles($conditionalStyles);
//conditions

//	Create the chart
$chart19= new PHPExcel_Chart(
	'chart19',		// name
	$title19,		// title
	$legend19,		// legend
	$plotarea19,		// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel19	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart19->setTopLeftPosition('I256');
$chart19->setBottomRightPosition('O266');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart19);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////


$objWorksheet->fromArray(
$results24, NULL, 'C275'
);

$dataseriesLabels20 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C274', 'Cash from Investing Activities');
$objPHPExcel->getActiveSheet()->getStyle("C274")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues20 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$275:$G275', NULL, 5),//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues20 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$276:$G276', NULL, 5));

//	Build the dataseries
$series20= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues20)-1),			// plotOrder
	$dataseriesLabels20,								// plotLabel
	$xAxisTickValues20,								// plotCategory
	$dataSeriesValues20								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series20->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea20 = new PHPExcel_Chart_PlotArea(NULL, array($series20));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend20=NULL;
$title20 = new PHPExcel_Chart_Title('Cash from Investing Activities');
$yAxisLabel20= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P302:P303');
$objPHPExcel->getActiveSheet()->mergeCells('R302:R303');
$objPHPExcel->getActiveSheet()->setCellValue('P300', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R300', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P301', '=IF(G302>C302,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R301', '=IF(G302>F302,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P302', '=round((G302-C302)/ABS(C302),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R302', '=round((G302-F302)/ABS(F302),2)');
$objPHPExcel->getActiveSheet()->getStyle('P301')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R301')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P302')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P302')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R302')->setConditionalStyles($conditionalStyles);
//conditions


//	Create the chart
$chart20= new PHPExcel_Chart(
	'chart20',		// name
	$title20,		// title
	$legend20,		// legend
	$plotarea20,		// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel20	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart20->setTopLeftPosition('I270');
$chart20->setBottomRightPosition('O280');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart20);
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////


$objWorksheet->fromArray(
$results25, NULL, 'C289'
);

$dataseriesLabels21 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C288', 'Cash from Financing Activities');
$objPHPExcel->getActiveSheet()->getStyle("C288")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues21 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$289:$G289', NULL, 5),//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues21 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$290:$G290', NULL, 5));

//	Build the dataseries
$series21= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues21)-1),			// plotOrder
	$dataseriesLabels21,								// plotLabel
	$xAxisTickValues21,								// plotCategory
	$dataSeriesValues21								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series21->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea21 = new PHPExcel_Chart_PlotArea(NULL, array($series21));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend21=NULL;
$title21 = new PHPExcel_Chart_Title('Cash from Financing Activities');
$yAxisLabel21= NULL;

//	Create the chart
$chart21= new PHPExcel_Chart(
	'chart21',		// name
	$title21,		// title
	$legend21,		// legend
	$plotarea21,		// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel21	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart21->setTopLeftPosition('I284');
$chart21->setBottomRightPosition('O294');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart21);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////

$objWorksheet->fromArray(
$results26, NULL, 'C301'
);

$dataseriesLabels22 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C300', 'Net Change in Cash');
$objPHPExcel->getActiveSheet()->getStyle("C300")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues22 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$302:$G302', NULL, 5),//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues22 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$303:$G303', NULL, 5));

//	Build the dataseries
$series22= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues22)-1),			// plotOrder
	$dataseriesLabels22,								// plotLabel
	$xAxisTickValues22,								// plotCategory
	$dataSeriesValues22								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series22->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea22 = new PHPExcel_Chart_PlotArea(NULL, array($series21));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend22=NULL;
$title22 = new PHPExcel_Chart_Title('Net Change in Cash');
$yAxisLabel22= NULL;

//	Create the chart
$chart22= new PHPExcel_Chart(
	'chart22',		// name
	$title22,		// title
	$legend22,		// legend
	$plotarea22,		// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel22	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart22->setTopLeftPosition('I297');
$chart22->setBottomRightPosition('O307');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart22);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////


$objWorksheet->fromArray(
$results27, NULL, 'C35'
);

$dataseriesLabels23 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C34', 'Gross Profit %');
$objPHPExcel->getActiveSheet()->getStyle("C34")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues23 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$35:$G35', NULL, 5),//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues23 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$36:$G36', NULL, 5));

//	Build the dataseries
$series23= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues23)-1),			// plotOrder
	$dataseriesLabels23,								// plotLabel
	$xAxisTickValues23,								// plotCategory
	$dataSeriesValues23								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series23->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea23 = new PHPExcel_Chart_PlotArea(NULL, array($series23));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend23=NULL;
$title23 = new PHPExcel_Chart_Title('Gross Profit %');
$yAxisLabel23= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P36:P37');
$objPHPExcel->getActiveSheet()->mergeCells('R36:R37');
$objPHPExcel->getActiveSheet()->setCellValue('P34', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R34', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P35', '=IF(G36>C36,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R35', '=IF(G36>F36,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P36', '=round((G36-C36)/ABS(C36),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R36', '=round((G36-F36)/ABS(F36),2)');
$objPHPExcel->getActiveSheet()->getStyle('P35')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R35')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P36')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P36')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R36')->setConditionalStyles($conditionalStyles);
//conditions

//	Create the chart
$chart23= new PHPExcel_Chart(
	'chart23',		// name
	$title23,		// title
	$legend23,		// legend
	$plotarea23,		// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel23	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart23->setTopLeftPosition('I30');
$chart23->setBottomRightPosition('O40');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart23);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////

$objWorksheet->fromArray(
$results28, NULL, 'C57'
);

$dataseriesLabels24 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C56', 'Operating Income (ETBT) %');
$objPHPExcel->getActiveSheet()->getStyle("C56")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues24 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$57:$G57', NULL, 5),//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues24 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$58:$G58', NULL, 5));

//	Build the dataseries
$series24= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues24)-1),			// plotOrder
	$dataseriesLabels24,							// plotLabel
	$xAxisTickValues24,								// plotCategory
	$dataSeriesValues24								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series24->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea24 = new PHPExcel_Chart_PlotArea(NULL, array($series24));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend24=NULL;
$title24 = new PHPExcel_Chart_Title('Operating Income (ETBT) %');
$yAxisLabel24= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P58:P59');
$objPHPExcel->getActiveSheet()->mergeCells('R58:R59');
$objPHPExcel->getActiveSheet()->setCellValue('P56', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R56', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P57', '=IF(G57>C57,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R57', '=IF(G57>F57,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P58', '=round((G58-C58)/ABS(C58),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R58', '=round((G58-F58)/ABS(F58),2)');
$objPHPExcel->getActiveSheet()->getStyle('P58')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R58')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P58')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P58')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R58')->setConditionalStyles($conditionalStyles);
//conditions

//	Create the chart
$chart24= new PHPExcel_Chart(
	'chart24',		// name
	$title24,		// title
	$legend24,		// legend
	$plotarea24,		// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel24	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart24->setTopLeftPosition('I53');
$chart24->setBottomRightPosition('O63');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart24);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////


$objWorksheet->fromArray(
$results29, NULL, 'C77'
);

$dataseriesLabels25 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C76', 'Income before tax %');
$objPHPExcel->getActiveSheet()->getStyle("C76")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues25 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$77:$G77', NULL, 5),//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues25 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$78:$G78', NULL, 5));

//	Build the dataseries
$series25= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues25)-1),			// plotOrder
	$dataseriesLabels25,							// plotLabel
	$xAxisTickValues25,								// plotCategory
	$dataSeriesValues25								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series25->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea25 = new PHPExcel_Chart_PlotArea(NULL, array($series25));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend25=NULL;
$title25 = new PHPExcel_Chart_Title('Income Before Tax %');
$yAxisLabel25= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P79:P80');
$objPHPExcel->getActiveSheet()->mergeCells('R79:R80');
$objPHPExcel->getActiveSheet()->setCellValue('P77', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R77', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P78', '=IF(G78>C78,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R78', '=IF(G78>F78,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P79', '=round((G78-C78)/ABS(C78),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R79', '=round((G78-F78)/ABS(F78),2)');
$objPHPExcel->getActiveSheet()->getStyle('P78')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R78')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P79')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P79')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R79')->setConditionalStyles($conditionalStyles);
//conditions

//	Create the chart
$chart25= new PHPExcel_Chart(
	'chart25',		// name
	$title25,		// title
	$legend25,		// legend
	$plotarea25,	// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel25	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart25->setTopLeftPosition('I74');
$chart25->setBottomRightPosition('O84');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart25);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////



$objWorksheet->fromArray(
$results30, NULL, 'C98'
);

$dataseriesLabels26 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C97', 'Net Profit%');
$objPHPExcel->getActiveSheet()->getStyle("C97")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues26 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$77:$G77', NULL, 5),//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues26 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$78:$G78', NULL, 5));

//	Build the dataseries
$series26= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues26)-1),			// plotOrder
	$dataseriesLabels26,							// plotLabel
	$xAxisTickValues26,								// plotCategory
	$dataSeriesValues26								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series26->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea26 = new PHPExcel_Chart_PlotArea(NULL, array($series26));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend26=NULL;
$title26 = new PHPExcel_Chart_Title('Net Profit %');
$yAxisLabel26= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P99:P100');
$objPHPExcel->getActiveSheet()->mergeCells('R99:R100');
$objPHPExcel->getActiveSheet()->setCellValue('P97', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R97', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P98', '=IF(G99>C99,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R98', '=IF(G99>F99,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P99', '=round((G99-C99)/ABS(C99),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R99', '=round((G99-F99)/ABS(F99),2)');
$objPHPExcel->getActiveSheet()->getStyle('P98')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R98')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P99')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P99')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R99')->setConditionalStyles($conditionalStyles);
//conditions

//	Create the chart
$chart26= new PHPExcel_Chart(
	'chart26',		// name
	$title26,		// title
	$legend26,		// legend
	$plotarea26,	// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel26	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart26->setTopLeftPosition('I94');
$chart26->setBottomRightPosition('O104');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart26);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////

$objWorksheet->fromArray(
$results31, NULL, 'C319'
);

$dataseriesLabels27 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C318', 'Current Ratio');
$objPHPExcel->getActiveSheet()->getStyle("C318")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues27 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$319:$G319', NULL, 5),//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues27 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$320:$G320', NULL, 5));

//	Build the dataseries
$series27= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues27)-1),			// plotOrder
	$dataseriesLabels27,							// plotLabel
	$xAxisTickValues27,								// plotCategory
	$dataSeriesValues27								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series27->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea27 = new PHPExcel_Chart_PlotArea(NULL, array($series27));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend27=NULL;
$title27 = new PHPExcel_Chart_Title('Current Ratio');
$yAxisLabel27= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P320:P321');
$objPHPExcel->getActiveSheet()->mergeCells('R320:R321');
$objPHPExcel->getActiveSheet()->setCellValue('P318', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R318', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P319', '=IF(G320>C320,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R319', '=IF(G320>F320,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P320', '=(G320-C320)/ABS(C320)');
$objPHPExcel->getActiveSheet()->setCellValue('R320', '=(G320-F320)/ABS(F320)');
$objPHPExcel->getActiveSheet()->getStyle('P319')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R319')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P320')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P320')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R320')->setConditionalStyles($conditionalStyles);
//conditions

//	Create the chart
$chart27= new PHPExcel_Chart(
	'chart27',		// name
	$title27,		// title
	$legend27,		// legend
	$plotarea27,	// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel27	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart27->setTopLeftPosition('I315');
$chart27->setBottomRightPosition('O325');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart27);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////



$objWorksheet->fromArray(
$results32, NULL, 'C333'
);

$dataseriesLabels28 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C332', 'Quick Ratio');
$objPHPExcel->getActiveSheet()->getStyle("C332")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues28 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$333:$G333', NULL, 5),//	Q1 to Q4
);
	
//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues28 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$334:$G334', NULL, 5));

//	Build the dataseries
$series28= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues28)-1),			// plotOrder
	$dataseriesLabels28,							// plotLabel
	$xAxisTickValues28,								// plotCategory
	$dataSeriesValues28								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series28->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea28 = new PHPExcel_Chart_PlotArea(NULL, array($series28));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend28=NULL;
$title28 = new PHPExcel_Chart_Title('Quick Ratio');
$yAxisLabel28= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P334:P335');
$objPHPExcel->getActiveSheet()->mergeCells('R334:R335');
$objPHPExcel->getActiveSheet()->setCellValue('P332', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R332', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P333', '=IF(G334>C334,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R333', '=IF(G334>F334,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P334', '=(G334-C334)/ABS(C334)');
$objPHPExcel->getActiveSheet()->setCellValue('R334', '=(G334-F334)/ABS(F334)');
$objPHPExcel->getActiveSheet()->getStyle('P333')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R333')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P334')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P334')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R334')->setConditionalStyles($conditionalStyles);
//conditions

//	Create the chart
$chart28= new PHPExcel_Chart(
	'chart28',		// name
	$title28,		// title
	$legend28,		// legend
	$plotarea28,	// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel28	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart28->setTopLeftPosition('I329');
$chart28->setBottomRightPosition('O339');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart28);

	///////////////
	///////////////
	//////////////
	///////////////
	

$objWorksheet->fromArray(
$results33, NULL, 'C348'
);

$dataseriesLabels29 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C347', 'Assets Turnover');
$objPHPExcel->getActiveSheet()->getStyle("C347")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues29 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$348:$G348', NULL, 5),//	Q1 to Q4
);
	
//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues29 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$349:$G349', NULL, 5));

//	Build the dataseries
$series29= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues29)-1),			// plotOrder
	$dataseriesLabels29,							// plotLabel
	$xAxisTickValues29,								// plotCategory
	$dataSeriesValues29								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series29->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea29 = new PHPExcel_Chart_PlotArea(NULL, array($series29));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend29=NULL;
$title29 = new PHPExcel_Chart_Title('Assets Turnover');
$yAxisLabel29= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P349:P350');
$objPHPExcel->getActiveSheet()->mergeCells('R349:R350');
$objPHPExcel->getActiveSheet()->setCellValue('P347', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R347', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P348', '=IF(G349>C349,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R348', '=IF(G349>F349,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P349', '=(G349-C349)/ABS(C349)');
$objPHPExcel->getActiveSheet()->setCellValue('R349', '=(G349-F349)/ABS(F349)');
$objPHPExcel->getActiveSheet()->getStyle('P349')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R349')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P349')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P349')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R349')->setConditionalStyles($conditionalStyles);
//conditions

//	Create the chart
$chart29= new PHPExcel_Chart(
	'chart29',		// name
	$title29,		// title
	$legend29,		// legend
	$plotarea29,	// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel29	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart29->setTopLeftPosition('I343');
$chart29->setBottomRightPosition('O353');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart29);
/////////////
//////////////
///////////////
/////////////


$objWorksheet->fromArray(
$results34, NULL, 'C362'
);

$dataseriesLabels30 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C361', 'Inventory Turnover');
$objPHPExcel->getActiveSheet()->getStyle("C361")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues30 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$362:$G362', NULL, 5),//	Q1 to Q4
);
	
//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues30 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$363:$G363', NULL, 5));

//	Build the dataseries
$series30 = new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues30)-1),			// plotOrder
	$dataseriesLabels30,							// plotLabel
	$xAxisTickValues30,								// plotCategory
	$dataSeriesValues30								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series30->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea30 = new PHPExcel_Chart_PlotArea(NULL, array($series30));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend30=NULL;
$title30 = new PHPExcel_Chart_Title('Inventory Turnover');
$yAxisLabel30= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P363:P364');
$objPHPExcel->getActiveSheet()->mergeCells('R363:R364');
$objPHPExcel->getActiveSheet()->setCellValue('P361', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R361', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P362', '=IF(G363>C363,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R362', '=IF(G363>F363,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P363', '=(G363-C363)/ABS(C363)');
$objPHPExcel->getActiveSheet()->setCellValue('R363', '=(G363-F363)/ABS(F363)');
$objPHPExcel->getActiveSheet()->getStyle('P362')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R362')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P363')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P363')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R363')->setConditionalStyles($conditionalStyles);
//conditions

//	Create the chart
$chart30= new PHPExcel_Chart(
	'chart30',		// name
	$title30,		// title
	$legend30,		// legend
	$plotarea30,	// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel30	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart30->setTopLeftPosition('I357');
$chart30->setBottomRightPosition('O367');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart30);
/////////
/////////
/////////
//////////



$objWorksheet->fromArray(
$results35, NULL, 'C376'
);

$dataseriesLabels31 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C375', 'Accounts Receivable Turnover');
$objPHPExcel->getActiveSheet()->getStyle("C375")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues31 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$376:$G376', NULL, 5),//	Q1 to Q4
);
	
//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues31 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$377:$G377', NULL, 5));

//	Build the dataseries
$series31 = new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues31)-1),			// plotOrder
	$dataseriesLabels31,							// plotLabel
	$xAxisTickValues31,								// plotCategory
	$dataSeriesValues31								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series31->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea31 = new PHPExcel_Chart_PlotArea(NULL, array($series31));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend31=NULL;
$title31 = new PHPExcel_Chart_Title('Accounts Receivable Turnover');
$yAxisLabel31= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P377:P378');
$objPHPExcel->getActiveSheet()->mergeCells('R377:R378');
$objPHPExcel->getActiveSheet()->setCellValue('P375', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R375', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P376', '=IF(G377>C377,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R376', '=IF(G377>F377,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P377', '=(G377-C377)/ABS(C377)');
$objPHPExcel->getActiveSheet()->setCellValue('R377', '=(G377-F377)/ABS(F377)');
$objPHPExcel->getActiveSheet()->getStyle('P376')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R376')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P377')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P377')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R377')->setConditionalStyles($conditionalStyles);
//conditions

//	Create the chart
$chart31= new PHPExcel_Chart(
	'chart31',		// name
	$title31,		// title
	$legend31,		// legend
	$plotarea31,	// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel31	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart31->setTopLeftPosition('I371');
$chart31->setBottomRightPosition('O381');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart31);
/////////
/////////
/////////
//////////

$objWorksheet->fromArray(
$results36, NULL, 'C390'
);

$dataseriesLabels32 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C389', 'Average Collection Period (Days)');
$objPHPExcel->getActiveSheet()->getStyle("C389")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues32 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$390:$G390', NULL, 5),//	Q1 to Q4
);
	
//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues32 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$391:$G391', NULL, 5));

//	Build the dataseries
$series32 = new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues31)-1),			// plotOrder
	$dataseriesLabels31,							// plotLabel
	$xAxisTickValues31,								// plotCategory
	$dataSeriesValues31								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series32->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea32 = new PHPExcel_Chart_PlotArea(NULL, array($series31));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend32=NULL;
$title32 = new PHPExcel_Chart_Title('Average Collection Period Days');
$yAxisLabel32= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P391:P392');
$objPHPExcel->getActiveSheet()->mergeCells('R391:R392');
$objPHPExcel->getActiveSheet()->setCellValue('P389', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R389', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P390', '=IF(G391>C391,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R390', '=IF(G391>F391,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P391', '=(G391-C391)/ABS(C391)');
$objPHPExcel->getActiveSheet()->setCellValue('R391', '=(G391-F391)/ABS(F391)');
$objPHPExcel->getActiveSheet()->getStyle('P390')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R390')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P391')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P391')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R391')->setConditionalStyles($conditionalStyles);
//conditions

//	Create the chart
$chart32= new PHPExcel_Chart(
	'chart32',		// name
	$title32,		// title
	$legend32,		// legend
	$plotarea32,	// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel32	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart32->setTopLeftPosition('I385');
$chart32->setBottomRightPosition('O395');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart32);
/////////
/////////
/////////
//////////


$objWorksheet->fromArray(
$results37, NULL, 'C404'
);

$dataseriesLabels33 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C403', 'Inventory Period (Days)');
$objPHPExcel->getActiveSheet()->getStyle("C403")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues33 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$404:$G404', NULL, 5),//	Q1 to Q4
);
	
//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues33 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$405:$G405', NULL, 5));

//	Build the dataseries
$series33 = new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues33)-1),			// plotOrder
	$dataseriesLabels33,							// plotLabel
	$xAxisTickValues33,								// plotCategory
	$dataSeriesValues33								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series33->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea33 = new PHPExcel_Chart_PlotArea(NULL, array($series33));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend33=NULL;
$title33 = new PHPExcel_Chart_Title('Inventory Period (Days)');
$yAxisLabel33= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P405:P406');
$objPHPExcel->getActiveSheet()->mergeCells('R405:R406');
$objPHPExcel->getActiveSheet()->setCellValue('P403', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R403', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P404', '=IF(G405>C405,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R404', '=IF(G405>F405,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P405', '=(G405-C405)/ABS(C405)');
$objPHPExcel->getActiveSheet()->setCellValue('R405', '=(G405-F405)/ABS(F405)');
$objPHPExcel->getActiveSheet()->getStyle('P404')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R404')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P405')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P405')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R405')->setConditionalStyles($conditionalStyles);
//conditions

//	Create the chart
$chart33= new PHPExcel_Chart(
	'chart33',		// name
	$title33,		// title
	$legend33,		// legend
	$plotarea33,	// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel33	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart33->setTopLeftPosition('I399');
$chart33->setBottomRightPosition('O409');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart33);
/////////
/////////
/////////
//////////


$objWorksheet->fromArray(
$results38, NULL, 'C418'
);

$dataseriesLabels34 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C417', 'Debt Ratio');
$objPHPExcel->getActiveSheet()->getStyle("C417")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues34 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$418:$G418', NULL, 5),//	Q1 to Q4
);
	
//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues34 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$419:$G419', NULL, 5));

//	Build the dataseries
$series34 = new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues34)-1),			// plotOrder
	$dataseriesLabels34,							// plotLabel
	$xAxisTickValues34,								// plotCategory
	$dataSeriesValues34								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series34->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea34 = new PHPExcel_Chart_PlotArea(NULL, array($series34));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend34=NULL;
$title34 = new PHPExcel_Chart_Title('Debt Ratio');
$yAxisLabel34= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P419:P420');
$objPHPExcel->getActiveSheet()->mergeCells('R419:R420');
$objPHPExcel->getActiveSheet()->setCellValue('P417', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R417', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P418', '=IF(G419>C419,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R418', '=IF(G419>F419,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P419', '=(G419-C419)/ABS(C419)');
$objPHPExcel->getActiveSheet()->setCellValue('R419', '=(G419-F419)/ABS(F419)');
$objPHPExcel->getActiveSheet()->getStyle('P418')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R418')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P419')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P419')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R419')->setConditionalStyles($conditionalStyles);
//conditions

//	Create the chart
$chart34= new PHPExcel_Chart(
	'chart34',		// name
	$title34,		// title
	$legend34,		// legend
	$plotarea34,	// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel34	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart34->setTopLeftPosition('I413');
$chart34->setBottomRightPosition('O423');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart34);
/////////
/////////
/////////
//////////


$objWorksheet->fromArray(
$results39, NULL, 'C431'
);

$dataseriesLabels35 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C430', 'Debt To Equity Ratio');
$objPHPExcel->getActiveSheet()->getStyle("C430")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues35 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$431:$G431', NULL, 5),//	Q1 to Q4
);
	
//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues35 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$432:$G432', NULL, 5));

//	Build the dataseries
$series35 = new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues35)-1),			// plotOrder
	$dataseriesLabels35,							// plotLabel
	$xAxisTickValues35,								// plotCategory
	$dataSeriesValues35								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series35->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea35 = new PHPExcel_Chart_PlotArea(NULL, array($series35));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend35=NULL;
$title35 = new PHPExcel_Chart_Title('Debt To Equity Ratio');
$yAxisLabel35= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P432:P433');
$objPHPExcel->getActiveSheet()->mergeCells('R432:R433');
$objPHPExcel->getActiveSheet()->setCellValue('P430', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R430', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P431', '=IF(G432>C432,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R431', '=IF(G432>F432,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P432', '=(G432-C432)/ABS(C432)');
$objPHPExcel->getActiveSheet()->setCellValue('R432', '=(G432-F432)/ABS(F432)');
$objPHPExcel->getActiveSheet()->getStyle('P431')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R431')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P432')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P432')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R432')->setConditionalStyles($conditionalStyles);
//conditions

//	Create the chart
$chart35= new PHPExcel_Chart(
	'chart35',		// name
	$title35,		// title
	$legend35,		// legend
	$plotarea35,	// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel35	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart35->setTopLeftPosition('I427');
$chart35->setBottomRightPosition('O437');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart35);
/////////
/////////
/////////
//////////


$objWorksheet->fromArray(
$results40, NULL, 'C446'
);

$dataseriesLabels36 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C445', 'Return on Assets % ');
$objPHPExcel->getActiveSheet()->getStyle("C445")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues36 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$446:$G446', NULL, 5),//	Q1 to Q4
);
	
//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues36 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$447:$G447', NULL, 5));

//	Build the dataseries
$series36 = new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues36)-1),			// plotOrder
	$dataseriesLabels36,							// plotLabel
	$xAxisTickValues36,								// plotCategory
	$dataSeriesValues36								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series36->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea36 = new PHPExcel_Chart_PlotArea(NULL, array($series36));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend36=NULL;
$title36 = new PHPExcel_Chart_Title('Return on Assets %');
$yAxisLabel36= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P447:P448');
$objPHPExcel->getActiveSheet()->mergeCells('R447:R448');
$objPHPExcel->getActiveSheet()->setCellValue('P445', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R445', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P446', '=IF(G447>C447,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R446', '=IF(G447>F447,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P447', '=(G447-C447)/ABS(C447)');
$objPHPExcel->getActiveSheet()->setCellValue('R447', '=(G447-F447)/ABS(F447)');
$objPHPExcel->getActiveSheet()->getStyle('P446')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R446')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P447')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P447')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R447')->setConditionalStyles($conditionalStyles);
//conditions

//	Create the chart
$chart36= new PHPExcel_Chart(
	'chart36',		// name
	$title36,		// title
	$legend36,		// legend
	$plotarea36,	// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel36	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart36->setTopLeftPosition('I441');
$chart36->setBottomRightPosition('O451');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart36);
/////////
/////////
/////////
//////////


$objWorksheet->fromArray(
$results42, NULL, 'C460'
);

$dataseriesLabels37 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C459', 'Return on Equity %');
$objPHPExcel->getActiveSheet()->getStyle("C459")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues37 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$460:$G460', NULL, 5),//	Q1 to Q4
);
	
//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues37 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$461:$G461', NULL, 5));

//	Build the dataseries
$series37 = new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues37)-1),			// plotOrder
	$dataseriesLabels37,							// plotLabel
	$xAxisTickValues37,								// plotCategory
	$dataSeriesValues37								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series37->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea37 = new PHPExcel_Chart_PlotArea(NULL, array($series37));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend37=NULL;
$title37 = new PHPExcel_Chart_Title('Return on Equity %');
$yAxisLabel37= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P461:P462');
$objPHPExcel->getActiveSheet()->mergeCells('R461:R462');
$objPHPExcel->getActiveSheet()->setCellValue('P459', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R459', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P460', '=IF(G461>C461,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R460', '=IF(G461>F461,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P461', '=(G461-C461)/ABS(C461)');
$objPHPExcel->getActiveSheet()->setCellValue('R461', '=(G461-F461)/ABS(F461)');
$objPHPExcel->getActiveSheet()->getStyle('P460')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R460')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P461')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P461')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R461')->setConditionalStyles($conditionalStyles);
//conditions

//	Create the chart
$chart37= new PHPExcel_Chart(
	'chart37',		// name
	$title37,		// title
	$legend37,		// legend
	$plotarea37,	// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel37	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart37->setTopLeftPosition('I455');
$chart37->setBottomRightPosition('O464');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart37);
/////////
/////////
/////////
//////////


$objWorksheet->fromArray(
$results41, NULL, 'C473'
);

$dataseriesLabels38 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C472', 'Net Working Capital Ratio');
$objPHPExcel->getActiveSheet()->getStyle("C472")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues38 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$473:$G473', NULL, 5),//	Q1 to Q4
);
	
//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues38 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$474:$G474', NULL, 5));

//	Build the dataseries
$series38 = new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues38)-1),			// plotOrder
	$dataseriesLabels38,							// plotLabel
	$xAxisTickValues38,								// plotCategory
	$dataSeriesValues38								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series38->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea38 = new PHPExcel_Chart_PlotArea(NULL, array($series38));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend38=NULL;
$title38 = new PHPExcel_Chart_Title('Net Working Capital Ratio %');
$yAxisLabel38= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P474:P475');
$objPHPExcel->getActiveSheet()->mergeCells('R474:R475');
$objPHPExcel->getActiveSheet()->setCellValue('P472', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R472', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P473', '=IF(G474>C474,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R473', '=IF(G474>F474,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P474', '=(G474-C474)/ABS(C474)');
$objPHPExcel->getActiveSheet()->setCellValue('R474', '=(G474-F474)/ABS(F474)');
$objPHPExcel->getActiveSheet()->getStyle('P473')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R473')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P474')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P474')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R474')->setConditionalStyles($conditionalStyles);
//conditions

//	Create the chart
$chart38= new PHPExcel_Chart(
	'chart38',		// name
	$title38,		// title
	$legend38,		// legend
	$plotarea38,	// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel38	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart38->setTopLeftPosition('I468');
$chart38->setBottomRightPosition('O478');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart38);
//$objPHPExcel->getActiveSheet()->setTitle('Financial_Ratios');
/////////
/////////
/////////
//////////


$newSheet=new PHPExcel_Worksheet($objPHPExcel,'Du_Pont_Analysis');
$objPHPExcel->addSheet($newSheet,1);
$objWorksheet=$objPHPExcel->setActiveSheetIndex(1);  



$query1 = "select income_statement.finyear, income_statement.totalrevenue_sales,income_statement.costrevenue,
income_statement.netprofit, 
balance_sheet.totalassets, balance_sheet.inventory, balance_sheet.accountrecive, balance_sheet.totalequity from income_statement, 
balance_sheet where income_statement.company='".$companyname."' and income_statement.finyear=balance_sheet.finyear ";

$resultnew = mysql_query($query1, $conn);
//$resultnew = mysqli_query($conn,$query5);
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
	new PHPExcel_Chart_DataSeriesValues('Number', 'Du_Pont_Analysis!$I$8:$M$8', NULL, 5),	//	2010
	
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
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$I$8:$M$8', NULL, 5),//	Q1 to Q4
);	

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues1 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Du_Pont_Analysis!I$9:$M$9', NULL, 5),);

//	Build the dataseries


$series1= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues1)-1),			// plotOrder
	$dataseriesLabels1,							    // plotLabel
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
$yAxisLabel1= NULL;


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
$objPHPExcel->getActiveSheet()->getStyle('L7')->setConditionalStyles($conditionalStyles);


//conditions

//	Create the chart
$chart1= new PHPExcel_Chart(
	'chart1',		// name
	$title1,		// title
	$legend1,		// legend
	$plotarea1,		// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel1	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart1->setTopLeftPosition('I10');
$chart1->setBottomRightPosition('N16');

//	Add the chart to the Financial_Ratios
$objPHPExcel->getActiveSheet()->addChart($chart1);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////



$objWorksheet->fromArray(
$results2,NULL,'D23'
);

$dataseriesLabels2 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$5', NULL, 1),	//	2012
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
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$D$23:$H23', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues2 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Du_Pont_Analysis!$D$24:$H24', NULL, 5));

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

//	Set the position where the chart should appear in the Du_Pont_Analysis

$chart2->setTopLeftPosition('D25');
$chart2->setBottomRightPosition('I31');

//	Add the chart to the Du_Pont_Analysis
$objWorksheet->addChart($chart2);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////


$objWorksheet->fromArray(
$results3,NULL,'N23'
);

$dataseriesLabels3 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$5', NULL, 1),	//	2012
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
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$N$23:$R23', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues3 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Du_Pont_Analysis!$N$24:$R24', NULL, 5));

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

//	Set the position where the chart should appear in the Du_Pont_Analysis

$chart3->setTopLeftPosition('N25');
$chart3->setBottomRightPosition('S31');

//	Add the chart to the Du_Pont_Analysis
$objWorksheet->addChart($chart3);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////



$objWorksheet->fromArray(
$results4,NULL,'A38'
);

$dataseriesLabels4 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$5', NULL, 1),	//	2012
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
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$38:$E38', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues4 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Du_Pont_Analysis!$A$39:$E39', NULL, 5));

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

//	Set the position where the chart should appear in the Du_Pont_Analysis

$chart4->setTopLeftPosition('A40');
$chart4->setBottomRightPosition('F46');

//	Add the chart to the Du_Pont_Analysis
$objWorksheet->addChart($chart4);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////


$objWorksheet->fromArray(
$results5,NULL,'R38'
);

$dataseriesLabels5 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$5', NULL, 1),	//	2012
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
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$R$38:$V38', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues5 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Du_Pont_Analysis!$R$39:$V39', NULL, 5));

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

//	Set the position where the chart should appear in the Du_Pont_Analysis

$chart5->setTopLeftPosition('R40');
$chart5->setBottomRightPosition('W46');

//	Add the chart to the Du_Pont_Analysis
$objWorksheet->addChart($chart5);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////




$objWorksheet->fromArray(
$results6,NULL,'A54'
);

$dataseriesLabels6 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$5', NULL, 1),	//	2012
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
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$54:$E54', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues6 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Du_Pont_Analysis!$A$55:$E55', NULL, 5));

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

//	Set the position where the chart should appear in the Du_Pont_Analysis

$chart6->setTopLeftPosition('A56');
$chart6->setBottomRightPosition('F62');

//	Add the chart to the Du_Pont_Analysis
$objWorksheet->addChart($chart6);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////





$objWorksheet->fromArray(
$results7,NULL,'H54'
);

$dataseriesLabels7 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$5', NULL, 1),	//	2012
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
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$H$54:$L54', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues7 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Du_Pont_Analysis!$H$55:$L55', NULL, 5));

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

//	Set the position where the chart should appear in the Du_Pont_Analysis

$chart7->setTopLeftPosition('H56');
$chart7->setBottomRightPosition('M62');

//	Add the chart to the Du_Pont_Analysis
$objWorksheet->addChart($chart7);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////



$objWorksheet->fromArray(
$results8,NULL,'O54'
);

$dataseriesLabels8 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$5', NULL, 1),	//	2012
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
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$O$54:$S54', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues8 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Du_Pont_Analysis!$O$55:$S55', NULL, 5));

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

//	Set the position where the chart should appear in the Du_Pont_Analysis

$chart8->setTopLeftPosition('O56');
$chart8->setBottomRightPosition('T62');

//	Add the chart to the Du_Pont_Analysis
$objWorksheet->addChart($chart8);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////



$objWorksheet->fromArray(
$results9,NULL,'V54'
);

$dataseriesLabels9 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$5', NULL, 1),	//	2012
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
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$V$54:$AA54', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues9 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Du_Pont_Analysis!$V$55:$AA55', NULL, 5));

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

//	Set the position where the chart should appear in the Du_Pont_Analysis

$chart9->setTopLeftPosition('V56');
$chart9->setBottomRightPosition('AA62');

//	Add the chart to the Du_Pont_Analysis
$objWorksheet->addChart($chart9);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////



$newSheet=new PHPExcel_Worksheet($objPHPExcel,'Financial_Dashboard');
$objPHPExcel->addSheet($newSheet,2);
$objWorksheet=$objPHPExcel->setActiveSheetIndex(2);  
$objPHPExcel->getActiveSheet();

$query2 = "select finyear, totalrevenue_sales, costrevenue,
grossprofit, operatingexp, operatinginc, 
incomebeforetax, netprofit
from income_statement where company='".$companyname."'";
$result = mysql_query($query2,$conn);

//$result = mysqli_query($conn,$query2);

$i=0;
$j=0;
$results1 = array();
$results2 = array();
$results3 = array();
$results4 = array();
$results5 = array();
$results6 = array();
$results7 = array();


while ($info = mysql_fetch_array($result,MYSQL_ASSOC))
{
	$finyear=$info['finyear'];
	$revsales=$info['totalrevenue_sales'];
	$costrevenue=$info['costrevenue'];
	$grossprofit=$info['grossprofit'];
	$operatingexp=$info['operatingexp'];
	$operatinginc=$info['operatinginc'];
	$incomebeforetax=$info['incomebeforetax'];
	$netprofit=$info['netprofit'];
	
	
	$results1[$i][$j]=$finyear;
	$results2[$i][$j]=$finyear;
	$results3[$i][$j]=$finyear;
	$results4[$i][$j]=$finyear;
	$results5[$i][$j]=$finyear;
	$results6[$i][$j]=$finyear;
	$results7[$i][$j]=$finyear;
	
	$j=$j+1;	
	
	$results1[$i+1][$j]=$revsales;
	$results2[$i+1][$j]=$costrevenue;
	$results3[$i+1][$j]=$grossprofit;
	$results4[$i+1][$j]=$operatingexp;
	$results5[$i+1][$j]=$operatinginc;
	$results6[$i+1][$j]=$incomebeforetax;
	$results7[$i+1][$j]=$netprofit;
}


$query3 = "select finyear, cashandequivalents, accountrecive, inventory,
totalcurrenassets, propertyequip, longterminvst,totallongterm, totalassets,
accountpayable, totalcurliab, longtermdebt, longtermliab, totalliabilities,
totalequity, liabilitiesandequity  
from balance_sheet where company='".$companyname."'";

$resultbalance = mysql_query($query3,$conn);

$results8 = array();
$results9 = array();
$results10 = array();
$results11 = array();
$results12 = array();
$results13 = array();
$results14 = array();
$results15 = array();
$results16 = array();
$results17 = array();
$results18 = array();
$results19 = array();
$results20 = array();
$results21 = array();

while ($info1 = mysql_fetch_array($resultbalance,MYSQL_ASSOC))
{
	$finyear = $info1['finyear'];
	$cashandequivalents = $info1['cashandequivalents'];
	$accountreceive = $info1['accountrecive'];
	$inventory = $info1['inventory'];
	$totalcurrentassets= $info1['totalcurrenassets'];
	$propertyequip= $info1['propertyequip'];
	$longterminvst= $info1['longterminvst'];
	$totallongterm= $info1['totallongterm'];
	$totalassets= $info1['totalassets'];
	$accountpayable= $info1['accountpayable'];
	$totalcurliab = $info1['totalcurliab'];
	$longtermdebt = $info1['longtermdebt'];
	$totalequity = $info1['totalequity'];
	$totalliabilities= $info1['totalliabilities'];
	$totalequity=$info1['totalequity'];
	$liabilitiesandequity=$info1['liabilitiesandequity'];  
	
	
	$results8[$i][$j]=$finyear;
	$results9[$i][$j]=$finyear;
	$results10[$i][$j]=$finyear;
	$results11[$i][$j]=$finyear;
	$results12[$i][$j]=$finyear;
	$results13[$i][$j]=$finyear;
	$results14[$i][$j]=$finyear;
	$results15[$i][$j]=$finyear;
	$results16[$i][$j]=$finyear;
	$results17[$i][$j]=$finyear;
	$results18[$i][$j]=$finyear;
	$results19[$i][$j]=$finyear;
	$results20[$i][$j]=$finyear;
	$results21[$i][$j]=$finyear;
	$results22[$i][$j]=$finyear;
	
	$j=$j+1;	
	
	$results8[$i+1][$j]=$cashandequivalents;
	$results9[$i+1][$j]=$accountreceive;
	$results10[$i+1][$j]=$inventory;
	$results11[$i+1][$j]=$totalcurrentassets;
	$results12[$i+1][$j]=$propertyequip;
	$results13[$i+1][$j]= $longterminvst;
	$results14[$i+1][$j]= $totallongterm;
	$results15[$i+1][$j]= $totalassets;
	$results16[$i+1][$j]= $accountpayable;
	$results17[$i+1][$j]= $totalcurliab;
	$results18[$i+1][$j]= $longtermdebt;
	$results19[$i+1][$j]= $totalequity;
	$results20[$i+1][$j]= $totalliabilities;
	$results21[$i+1][$j]= $totalequity;
	$results22[$i+1][$j]= $liabilitiesandequity;
}
  
$query4 = "select finyear, cashfromoper, cashfrominvesting, cashfromfinance
from cash_flow where company='".$companyname."'";

$resultcash = mysql_query($query4, $conn);

$results23 = array();
$results24 = array();
$results25 = array();

while ($info4 = mysql_fetch_array($resultcash,MYSQL_ASSOC))
{
	$finyear = $info4['finyear'];
	$cashfromoper = $info4['cashfromoper'];
	$cashfrominvesting = $info4['cashfrominvesting'];
	$cashfromfinance = $info4['cashfromfinance'];
	
	$results23[$i][$j]=$finyear;
	$results24[$i][$j]=$finyear;
	$results25[$i][$j]=$finyear;
	
	$j=$j+1;	
	
	$results23[$i+1][$j]=$cashfromoper;
	$results24[$i+1][$j]=$cashfrominvesting;
	$results25[$i+1][$j]=$cashfromfinance;
}
  
  

$sheet = $objPHPExcel->getActiveSheet();
$sheet->getStyle('A1:AC1')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->getStyle('A1:AC1')->getFill()->getStartColor()->setRGB('008000');

$objPHPExcel->getActiveSheet()->mergeCells('A1:AC1');
$objPHPExcel->getActiveSheet()->setCellValue('A1', 'Financial DashBoard');
$objPHPExcel->getActiveSheet()->getStyle("A1")->getFont()->setBold(true)
                                ->setName('Verdana')
                                ->setSize(20)
								->getColor()->setRGB('FFFFFF');

$objPHPExcel->getActiveSheet()->getStyle('A1:AC1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);


$objWorksheet->fromArray(
$results1,NULL,'A4'
);

//$objWorksheet->fromArray(
//	array(
//		array('',	2010,	2011,	2012),
//		array('Q1',   12,   15,		21),
//		array('Q2',   56,   73,		86),
//		array('Q3',   52,   61,		69),
//		array('Q4',   30,   32,		0),
//	)
//);

//	Set the Labels for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataseriesLabels1 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$A$4:$E$4', NULL, 5),
	
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$xAxisTickValues1 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$4:$E$4', NULL, 5),
	);	//	Q1 to Q4);
//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues1 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$A$5:E$5', NULL, 5),
	);

//	Build the dataseries
$series1 = new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues1)-1),			// plotOrder
	$dataseriesLabels1,								// plotLabel
	$xAxisTickValues1,								// plotCategory
	$dataSeriesValues1								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series1->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea1 = new PHPExcel_Chart_PlotArea(NULL, array($series1));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend1=NULL;
$title1 = new PHPExcel_Chart_Title('Total Revenue Sales');
$yAxisLabel1 = NULL;



//	Create the chart
$chart1 = new PHPExcel_Chart(
	'chart1',		// name
	$title1,			// title
	$legend1,		// legend
	$plotarea1,		// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel1	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Dashboard
$chart1->setTopLeftPosition('A4');
$chart1->setBottomRightPosition('F21');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart1);
 
 //////////////////
 
//////////////////
 
//////////////////
 
//////////////////
 
//////////////////
 

$objPHPExcel->getActiveSheet()->fromArray($results2, NULL, 'G4');


$dataseriesLabels2 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$G$4', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$G$5', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$G$6', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$G$7', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$G$8', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues2 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$G$4:$K$4', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues2 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$G$5:K$5', NULL, 5));

//	Build the dataseries
$series2 = new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues2)-1),			// plotOrder
	$dataseriesLabels2,								// plotLabel
	$xAxisTickValues2,								// plotCategory
	$dataSeriesValues2								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series2->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea2 = new PHPExcel_Chart_PlotArea(NULL, array($series2));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend2=NULL;
$title2 = new PHPExcel_Chart_Title('Cost of Revenue/COGS');
$yAxisLabel2 = NULL;

//	Create the chart
$chart2 = new PHPExcel_Chart(
	'chart2',		// name
	$title2,			// title
	$legend2,		// legend
	$plotarea2,		// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel2		// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Dashboard
$chart2->setTopLeftPosition('G4');
$chart2->setBottomRightPosition('L20');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart2);
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
$objPHPExcel->getActiveSheet()->fromArray($results3, NULL, 'M4');

$dataseriesLabels3 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$M$4', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$M$5', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$M$6', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$M$7', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$M$8', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues3 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$M$4:$Q$4', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues3 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$M$5:Q$5', NULL, 5));

//	Build the dataseries
$series3 = new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues3)-1),			// plotOrder
	$dataseriesLabels3,								// plotLabel
	$xAxisTickValues3,								// plotCategory
	$dataSeriesValues3								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series3->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea3 = new PHPExcel_Chart_PlotArea(NULL, array($series3));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend3=NULL;
$title3 = new PHPExcel_Chart_Title('Gross Profit');
$yAxisLabel3 = NULL;

//	Create the chart
$chart3 = new PHPExcel_Chart(
	'chart3',		// name
	$title3,		// title
	$legend3,		// legend
	$plotarea3,		// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel3	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Dashboard
$chart3->setTopLeftPosition('M4');
$chart3->setBottomRightPosition('R20');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart3);
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
$objPHPExcel->getActiveSheet()->fromArray($results4, NULL, 'S4');


$dataseriesLabels4 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$S$4', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$S$5', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$S$6', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$S$7', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$S$8', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues4 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$S$4:$W$4', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues4 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$S$5:W$5', NULL, 5));

//	Build the dataseries
$series4 = new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues4)-1),			// plotOrder
	$dataseriesLabels4,								// plotLabel
	$xAxisTickValues4,								// plotCategory
	$dataSeriesValues4								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series4->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea4 = new PHPExcel_Chart_PlotArea(NULL, array($series4));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend4=NULL;
$title4 = new PHPExcel_Chart_Title('Operating Expenses');
$yAxisLabel4= NULL;

//	Create the chart
$chart4 = new PHPExcel_Chart(
	'chart4',		// name
	$title4,		// title
	$legend4,		// legend
	$plotarea4,		// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel4	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Dashboard
$chart4->setTopLeftPosition('S4');
$chart4->setBottomRightPosition('X20');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart4);
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
$objPHPExcel->getActiveSheet()->fromArray($results5, NULL, 'Y4');


$dataseriesLabels5 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$Y$4', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$Y$5', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$Y$6', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$Y$6', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$Y$6', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues5 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$Y$4:$AC$4', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues5 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$Y$5:AC$5', NULL, 5));

//	Build the dataseries
$series5 = new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues5)-1),			// plotOrder
	$dataseriesLabels5,								// plotLabel
	$xAxisTickValues5,								// plotCategory
	$dataSeriesValues5								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series5->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea5 = new PHPExcel_Chart_PlotArea(NULL, array($series5));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend5=NULL;
$title5 = new PHPExcel_Chart_Title('Operating Income');
$yAxisLabel5= NULL;

//	Create the chart
$chart5 = new PHPExcel_Chart(
	'chart5',		// name
	$title5,		// title
	$legend5,		// legend
	$plotarea5,		// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel5	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Dashboard
$chart5->setTopLeftPosition('Y4');
$chart5->setBottomRightPosition('AD20');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart5);
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////

$objPHPExcel->getActiveSheet()->fromArray($results6, NULL, 'A23');


$dataseriesLabels6 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$23', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$24', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$25', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$26', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$27', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues6 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$23:$F$23', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues6 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$A$24:F$24', NULL, 5));

//	Build the dataseries
$series6 = new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues6)-1),			// plotOrder
	$dataseriesLabels6,								// plotLabel
	$xAxisTickValues6,								// plotCategory
	$dataSeriesValues6								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series6->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea6 = new PHPExcel_Chart_PlotArea(NULL, array($series6));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend6=NULL;
$title6 = new PHPExcel_Chart_Title('Income Before Tax');
$yAxisLabel6= NULL;

//	Create the chart
$chart6 = new PHPExcel_Chart(
	'chart6',		// name
	$title6,		// title
	$legend6,		// legend
	$plotarea6,		// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel6	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Dashboard
$chart6->setTopLeftPosition('A23');
$chart6->setBottomRightPosition('F40');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart6);
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
$objPHPExcel->getActiveSheet()->fromArray($results7, NULL, 'G23');


$dataseriesLabels7 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$G$23', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$G$24', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$G$25', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$G$26', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$G$27', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues7 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$G$23:$L$23', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues7 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$G$24:L$24', NULL, 5));

//	Build the dataseries
$series7 = new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues7)-1),			// plotOrder
	$dataseriesLabels7,								// plotLabel
	$xAxisTickValues7,								// plotCategory
	$dataSeriesValues7								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series7->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea7 = new PHPExcel_Chart_PlotArea(NULL, array($series7));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend7=NULL;
$title7 = new PHPExcel_Chart_Title('Net Profit');
$yAxisLabel7= NULL;

//	Create the chart
$chart7 = new PHPExcel_Chart(
	'chart7',		// name
	$title7,		// title
	$legend7,		// legend
	$plotarea7,		// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel7	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Dashboard
$chart7->setTopLeftPosition('G23');
$chart7->setBottomRightPosition('L40');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart7);
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
$objPHPExcel->getActiveSheet()->fromArray($results8, NULL, 'M23');

$dataseriesLabels8 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$M$23', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$M$24', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$M$25', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$M$26', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$M$27', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues8 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$M$23:$Q$23', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues8 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$M$24:Q$24', NULL, 5));

//	Build the dataseries
$series8 = new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues8)-1),			// plotOrder
	$dataseriesLabels8,								// plotLabel
	$xAxisTickValues8,								// plotCategory
	$dataSeriesValues8								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series8->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea8 = new PHPExcel_Chart_PlotArea(NULL, array($series8));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend8=NULL;
$title8 = new PHPExcel_Chart_Title('Cash & Equivalents');
$yAxisLabel8= NULL;

//	Create the chart
$chart8 = new PHPExcel_Chart(
	'chart8',		// name
	$title8,		// title
	$legend8,		// legend
	$plotarea8,		// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel8	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Dashboard
$chart8->setTopLeftPosition('M23');
$chart8->setBottomRightPosition('R40');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart8);
//////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////
$objPHPExcel->getActiveSheet()->fromArray($results9, NULL, 'S23');

$dataseriesLabels9 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues9 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$S$23:$W$23', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues9 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$S$24:W$24', NULL, 5));

//	Build the dataseries
$series9 = new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues9)-1),			// plotOrder
	$dataseriesLabels9,								// plotLabel
	$xAxisTickValues9,								// plotCategory
	$dataSeriesValues9								// plotValues
);

//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series9->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea9 = new PHPExcel_Chart_PlotArea(NULL, array($series9));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend9=NULL;
$title9 = new PHPExcel_Chart_Title('Account Receivables');
$yAxisLabel9= NULL;

//	Create the chart
$chart9 = new PHPExcel_Chart(
	'chart9',		// name
	$title9,		// title
	$legend9,		// legend
	$plotarea9,		// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel9	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Dashboard
$chart9->setTopLeftPosition('S23');
$chart9->setBottomRightPosition('X40');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart9);


/////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////
$objPHPExcel->getActiveSheet()->fromArray($results10, NULL, 'Y23');

$dataseriesLabels10 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues10 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$Y$23:$AC$23', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues10 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$Y$24:AC$24', NULL, 5));

//	Build the dataseries
$series10 = new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues10)-1),			// plotOrder
	$dataseriesLabels10,							// plotLabel
	$xAxisTickValues10,								// plotCategory
	$dataSeriesValues10								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series10->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea10 = new PHPExcel_Chart_PlotArea(NULL, array($series10));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend10=NULL;
$title10 = new PHPExcel_Chart_Title('Inventory');
$yAxisLabel10= NULL;

//	Create the chart
$chart10= new PHPExcel_Chart(
	'chart10',		// name
	$title10,		// title
	$legend10,		// legend
	$plotarea10,	// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel10	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Dashboard
$chart10->setTopLeftPosition('Y23');
$chart10->setBottomRightPosition('AD40');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart10);

//////////////////////////////
///////////////////////////////
//////////////////////////////
///////////////////////////////
//////////////////////////////
///////////////////////////////

$objPHPExcel->getActiveSheet()->fromArray($results11, NULL, 'A43');

$dataseriesLabels11 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues11 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$43:$E$43', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues11 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$A$44:$E$44', NULL, 5));

//	Build the dataseries
$series11 = new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues11)-1),			// plotOrder
	$dataseriesLabels11,							// plotLabel
	$xAxisTickValues11,								// plotCategory
	$dataSeriesValues11								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series11->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea11 = new PHPExcel_Chart_PlotArea(NULL, array($series11));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend11=NULL;
$title11 = new PHPExcel_Chart_Title('Total Current Assets');
$yAxisLabel11= NULL;

//	Create the chart
$chart11= new PHPExcel_Chart(
	'chart11',		// name
	$title11,		// title
	$legend11,		// legend
	$plotarea11,	// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel11	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Dashboard
$chart11->setTopLeftPosition('A43');
$chart11->setBottomRightPosition('F60');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart11);

//////////////////////////////
///////////////////////////////
//////////////////////////////
///////////////////////////////
//////////////////////////////
///////////////////////////////

$objPHPExcel->getActiveSheet()->fromArray($results12, NULL, 'G43');

$dataseriesLabels12 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues12 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$G$43:$K$43', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues12 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$G$44:$K44', NULL, 5));

//	Build the dataseries
$series12 = new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues12)-1),			// plotOrder
	$dataseriesLabels12,							// plotLabel
	$xAxisTickValues12,								// plotCategory
	$dataSeriesValues12								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series12->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea12 = new PHPExcel_Chart_PlotArea(NULL, array($series12));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend12=NULL;
$title12 = new PHPExcel_Chart_Title('Property/Equipments');
$yAxisLabel12= NULL;

//	Create the chart
$chart12= new PHPExcel_Chart(
	'chart12',		// name
	$title12,		// title
	$legend12,		// legend
	$plotarea12,	// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel2	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Dashboard
$chart12->setTopLeftPosition('G43');
$chart12->setBottomRightPosition('L60');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart12);

//////////////////////////////
///////////////////////////////
//////////////////////////////
///////////////////////////////
//////////////////////////////
///////////////////////////////


$objPHPExcel->getActiveSheet()->fromArray($results13, NULL, 'M43');

$dataseriesLabels13 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues13 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$M$43:$R$43', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues13 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$M$44:$R44', NULL, 5));

//	Build the dataseries
$series13 = new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues13)-1),			// plotOrder
	$dataseriesLabels13,							// plotLabel
	$xAxisTickValues13,								// plotCategory
	$dataSeriesValues13								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series13->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea13 = new PHPExcel_Chart_PlotArea(NULL, array($series13));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend13=NULL;
$title13 = new PHPExcel_Chart_Title('Long Term Investments');
$yAxisLabel13= NULL;

//	Create the chart
$chart13= new PHPExcel_Chart(
	'chart13',		// name
	$title13,		// title
	$legend13,		// legend
	$plotarea13,	// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel3	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Dashboard
$chart13->setTopLeftPosition('M43');
$chart13->setBottomRightPosition('R60');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart13);

////////////////////////////////////////////////////////
///////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////

$objPHPExcel->getActiveSheet()->fromArray($results14, NULL, 'S43');

$dataseriesLabels14 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues14 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$S$43:$W$43', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues14 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$S$44:$W44', NULL, 5));

//	Build the dataseries
$series14 = new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues14)-1),			// plotOrder
	$dataseriesLabels14,							// plotLabel
	$xAxisTickValues14,								// plotCategory
	$dataSeriesValues14								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series14->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea14 = new PHPExcel_Chart_PlotArea(NULL, array($series14));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend14=NULL;
$title14 = new PHPExcel_Chart_Title('Long Term Liabilities');
$yAxisLabel14= NULL;

//	Create the chart
$chart14= new PHPExcel_Chart(
	'chart14',		// name
	$title14,		// title
	$legend14,		// legend
	$plotarea14,	// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel4	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Dashboard
$chart14->setTopLeftPosition('S43');
$chart14->setBottomRightPosition('X60');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart14);

/////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////


$objPHPExcel->getActiveSheet()->fromArray($results15, NULL, 'Y43');

$dataseriesLabels15 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues15 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$Y$43:$AC$43', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues15 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$Y$44:$AC44', NULL, 5),);

//	Build the dataseries
$series15 = new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues15)-1),			// plotOrder
	$dataseriesLabels15,							// plotLabel
	$xAxisTickValues15,								// plotCategory
	$dataSeriesValues15								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series15->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea15 = new PHPExcel_Chart_PlotArea(NULL, array($series15));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend15=NULL;
$title15 = new PHPExcel_Chart_Title('Total Assets');
$yAxisLabel15= NULL;

//	Create the chart
$chart15= new PHPExcel_Chart(
	'chart15',		// name
	$title15,		// title
	$legend15,		// legend
	$plotarea15,	// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel5	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Dashboard
$chart15->setTopLeftPosition('Y43');
$chart15->setBottomRightPosition('AD60');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart15);

/////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////


$objPHPExcel->getActiveSheet()->fromArray($results16, NULL, 'A63');

$dataseriesLabels16 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues16 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$63:$E$63', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues16 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$A$64:E$64', NULL, 5));

//	Build the dataseries
$series16 = new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues16)-1),			// plotOrder
	$dataseriesLabels16,							// plotLabel
	$xAxisTickValues16,								// plotCategory
	$dataSeriesValues16								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series16->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea16 = new PHPExcel_Chart_PlotArea(NULL, array($series16));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend16=NULL;
$title16 = new PHPExcel_Chart_Title('Accounts Payable');
$yAxisLabel16= NULL;

//	Create the chart
$chart16= new PHPExcel_Chart(
	'chart16',		// name
	$title16,		// title
	$legend16,		// legend
	$plotarea16,	// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel16	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Dashboard
$chart16->setTopLeftPosition('A63');
$chart16->setBottomRightPosition('F80');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart16);

//////////////////////////////
///////////////////////////////
//////////////////////////////
///////////////////////////////
//////////////////////////////
///////////////////////////////

$objPHPExcel->getActiveSheet()->fromArray($results17, NULL, 'G63');

$dataseriesLabels17 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues17 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$G$63:$K$63', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues17 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$G$64:$K64', NULL, 5));

//	Build the dataseries
$series17 = new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues17)-1),			// plotOrder
	$dataseriesLabels17,							// plotLabel
	$xAxisTickValues17,								// plotCategory
	$dataSeriesValues17								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series17->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea17 = new PHPExcel_Chart_PlotArea(NULL, array($series17));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend17=NULL;
$title17 = new PHPExcel_Chart_Title('Total Current Liabilities');
$yAxisLabel17= NULL;

//	Create the chart
$chart17= new PHPExcel_Chart(
	'chart17',		// name
	$title17,		// title
	$legend17,		// legend
	$plotarea17,	// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel7	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Dashboard
$chart17->setTopLeftPosition('G63');
$chart17->setBottomRightPosition('L80');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart17);

//////////////////////////////
///////////////////////////////
//////////////////////////////
///////////////////////////////
//////////////////////////////
///////////////////////////////


$objPHPExcel->getActiveSheet()->fromArray($results18, NULL, 'M63');

$dataseriesLabels18 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues18 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$M$63:$R$63', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues18 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$M$64:$R64', NULL, 5));

//	Build the dataseries
$series18 = new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues18)-1),			// plotOrder
	$dataseriesLabels18,							// plotLabel
	$xAxisTickValues18,								// plotCategory
	$dataSeriesValues18								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series18->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea18 = new PHPExcel_Chart_PlotArea(NULL, array($series18));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend18=NULL;
$title18 = new PHPExcel_Chart_Title('Long Term Debt');
$yAxisLabel18= NULL;

//	Create the chart
$chart18= new PHPExcel_Chart(
	'chart18',		// name
	$title18,		// title
	$legend18,		// legend
	$plotarea18,	// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel8	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Dashboard
$chart18->setTopLeftPosition('M63');
$chart18->setBottomRightPosition('R80');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart18);

////////////////////////////////////////////////////////
///////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////



$objPHPExcel->getActiveSheet()->fromArray($results19, NULL, 'S63');

$dataseriesLabels19 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues19 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$S$63:$X$63', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues19 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$S$64:$X64', NULL, 5));

//	Build the dataseries
$series19 = new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues19)-1),			// plotOrder
	$dataseriesLabels19,							// plotLabel
	$xAxisTickValues19,								// plotCategory
	$dataSeriesValues19								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series19->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea19 = new PHPExcel_Chart_PlotArea(NULL, array($series19));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend19=NULL;
$title19 = new PHPExcel_Chart_Title('Long Term Liabilities');
$yAxisLabel19= NULL;

//	Create the chart
$chart19= new PHPExcel_Chart(
	'chart19',		// name
	$title19,		// title
	$legend19,		// legend
	$plotarea19,	// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel19	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Dashboard
$chart19->setTopLeftPosition('S63');
$chart19->setBottomRightPosition('X80');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart19);

////////////////////////////////////////////////////////
///////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////
$objPHPExcel->getActiveSheet()->fromArray($results20, NULL, 'Y63');

$dataseriesLabels20 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues20 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$Y$63:$AC63', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues20 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$Y$64:$AC64', NULL, 5));

//	Build the dataseries
$series20 = new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues20)-1),			// plotOrder
	$dataseriesLabels20,							// plotLabel
	$xAxisTickValues20,								// plotCategory
	$dataSeriesValues20								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series20->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea20 = new PHPExcel_Chart_PlotArea(NULL, array($series20));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend20=NULL;
$title20 = new PHPExcel_Chart_Title('Total Liabilities');
$yAxisLabel20= NULL;

//	Create the chart
$chart20= new PHPExcel_Chart(
	'chart20',		// name
	$title20,		// title
	$legend20,		// legend
	$plotarea20,	// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel20	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Dashboard
$chart20->setTopLeftPosition('Y63');
$chart20->setBottomRightPosition('AD80');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart20);

////////////////////////////////////////////////////////
///////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////

$objPHPExcel->getActiveSheet()->fromArray($results21, NULL, 'A83');

$dataseriesLabels21 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues21 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$83:$F$83', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues21 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$A$84:$F$84', NULL, 5));

//	Build the dataseries
$series21 = new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues21)-1),			// plotOrder
	$dataseriesLabels21,							// plotLabel
	$xAxisTickValues21,								// plotCategory
	$dataSeriesValues21								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series21->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea21 = new PHPExcel_Chart_PlotArea(NULL, array($series21));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend21=NULL;
$title21 = new PHPExcel_Chart_Title('Total Equity');
$yAxisLabel21= NULL;

//	Create the chart
$chart21= new PHPExcel_Chart(
	'chart21',		// name
	$title21,		// title
	$legend21,		// legend
	$plotarea21,	// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel21	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Dashboard
$chart21->setTopLeftPosition('A83');
$chart21->setBottomRightPosition('F100');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart21);

////////////////////////////////////////////////////////
///////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////
$objPHPExcel->getActiveSheet()->fromArray($results22, NULL, 'G83');

$dataseriesLabels22 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues22 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$G$83:$L83', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues22 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$G$84:$L84', NULL, 5));

//	Build the dataseries
$series22 = new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues22)-1),			// plotOrder
	$dataseriesLabels22,							// plotLabel
	$xAxisTickValues22,								// plotCategory
	$dataSeriesValues22								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series22->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea22 = new PHPExcel_Chart_PlotArea(NULL, array($series22));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend22=NULL;
$title22 = new PHPExcel_Chart_Title('Cash from Operations');
$yAxisLabel22= NULL;

//	Create the chart
$chart22= new PHPExcel_Chart(
	'chart22',		// name
	$title22,		// title
	$legend22,		// legend
	$plotarea22,	// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel22	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Dashboard
$chart22->setTopLeftPosition('G83');
$chart22->setBottomRightPosition('L100');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart22);

////////////////////////////////////////////////////////
///////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////

$objPHPExcel->getActiveSheet()->fromArray($results23, NULL, 'M83');

$dataseriesLabels23 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues23 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$M$83:$R83', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues23 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$M$84:$R84', NULL, 5));

//	Build the dataseries
$series23 = new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues23)-1),			// plotOrder
	$dataseriesLabels23,							// plotLabel
	$xAxisTickValues23,								// plotCategory
	$dataSeriesValues23								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series23->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea23 = new PHPExcel_Chart_PlotArea(NULL, array($series23));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend23=NULL;
$title23 = new PHPExcel_Chart_Title('Liabilities & Equity');
$yAxisLabel23= NULL;

//	Create the chart
$chart23= new PHPExcel_Chart(
	'chart23',		// name
	$title23,		// title
	$legend23,		// legend
	$plotarea23,	// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel23	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Dashboard
$chart23->setTopLeftPosition('M83');
$chart23->setBottomRightPosition('R100');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart23);

////////////////////////////////////////////////////////
///////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////

$objPHPExcel->getActiveSheet()->fromArray($results24, NULL, 'S83');

$dataseriesLabels24 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues24 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$S$83:$W83', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues24 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$S$84:$W84', NULL, 5));

//	Build the dataseries
$series24 = new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues24)-1),			// plotOrder
	$dataseriesLabels24,							// plotLabel
	$xAxisTickValues24,								// plotCategory
	$dataSeriesValues24								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series24->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea24 = new PHPExcel_Chart_PlotArea(NULL, array($series24));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend24=NULL;
$title24 = new PHPExcel_Chart_Title('Cash from Investing');
$yAxisLabel24= NULL;

//	Create the chart
$chart24= new PHPExcel_Chart(
	'chart24',		// name
	$title24,		// title
	$legend24,		// legend
	$plotarea24,	// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel24	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Dashboard
$chart24->setTopLeftPosition('S83');
$chart24->setBottomRightPosition('X100');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart24);

////////////////////////////////////////////////////////
///////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////

$objPHPExcel->getActiveSheet()->fromArray($results25, NULL, 'Y83');

$dataseriesLabels25 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues25 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$Y$83:$AC83', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues25 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$Y$84:$AC84', NULL, 5));

//	Build the dataseries
$series25= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues25)-1),			// plotOrder
	$dataseriesLabels25,							// plotLabel
	$xAxisTickValues25,								// plotCategory
	$dataSeriesValues25								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series25->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea25 = new PHPExcel_Chart_PlotArea(NULL, array($series25));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend25=NULL;
$title25 = new PHPExcel_Chart_Title('Cash From Financing');
$yAxisLabel25= NULL;

//	Create the chart
$chart25= new PHPExcel_Chart(
	'chart25',		// name
	$title25,		// title
	$legend25,		// legend
	$plotarea25,	// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel25	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Dashboard
$chart25->setTopLeftPosition('Y83');
$chart25->setBottomRightPosition('AD100');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart25);


////////////////////////////////////////////////////////
///////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////


// Save Excel 2007 file
echo date('H:i:s') , " Write to Excel2007 format" , EOL;
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->setIncludeCharts(TRUE);
$file=dirname(__FILE__).'/outputfilefolder/output'.$companyname;
$file=$file.".xlsx";
$objWriter->save($file);
echo date('H:i:s') , " File written to " , str_replace('.php', '.xlsx', pathinfo($file, PATHINFO_BASENAME)) , EOL;
$queryuploadfile="update uploadfiledata set outputfilename='output".$companyname.".xlsx' where companyname='".$companyname."'"; 
mysql_query($queryuploadfile,$conn);
// Echo memory peak usage
echo date('H:i:s') , " Peak memory usage: " , (memory_get_peak_usage(true) / 1024 / 1024) , " MB" , EOL;
// Echo done
echo date('H:i:s') , " Done writing file" , EOL;
echo 'File has been created in ' , getcwd() , EOL;
header('location:convertfile.php');
ob_end_flush();
?>