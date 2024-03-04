<?php
session_start();
ob_start();

$companyname =$_SESSION['companyname'];
$inputfilename =$_SESSION['inputfilename'];

if ($_GET['companyname'] !="" )
{
$companyname=$_GET['companyname'];
}

if ($_GET['inputfile'] !="" )
{
$inputfilename=$_GET['inputfile'];
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
$conn=mysqli_connect('sql211.epizy.com','epiz_33239466','Horizon2022','epiz_33239466_new','3306','') or  die ("Error in query:".mysqli_error());
	
$query1 = "select finyear, totalrevenue_sales, costrevenue,
grossprofit, operatingexp, operatinginc, 
incomebeforetax, netprofit 
from income_statement where company='".$companyname."'";
$result = mysqli_query($conn,$query1);

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

while ($info = mysqli_fetch_array($result,MYSQLI_ASSOC))
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

$resultbalance = mysqli_query($conn,$query3);

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

while ($info1 = mysqli_fetch_array($resultbalance,MYSQLI_ASSOC))
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

$resultcash = mysqli_query($conn,$query4);
//$resultcash = mysqli_query($conn,$query4);

$results23 = array();
$results24 = array();
$results25 = array();
$results26 = array();

//while ($info4 = mysqli_fetch_array($resultcash,MYSQLI_ASSOC))

while ($info4 = mysqli_fetch_array($resultcash,MYSQLI_ASSOC))
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
balance_sheet where income_statement.company='" . $companyname . "' and 
income_statement.finyear=balance_sheet.finyear and income_statement.company=balance_sheet.company";


$resultnew = mysqli_query($conn,$query5);
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
while ($info5 = mysqli_fetch_array($resultnew,MYSQLI_ASSOC))
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
	$results40[$i+1][$j]=round($netprofit/$totalassets*100);
	$results42[$i+1][$j]=round($netprofit/$totalequity*100,2);
	}
 
$objPHPExcel = new PHPExcel();

$newSheet=new PHPExcel_Worksheet($objPHPExcel,'Financial_Ratios');
$objPHPExcel->addSheet($newSheet,0);
$objWorksheet=$objPHPExcel->setActiveSheetIndex(0);  
$objPHPExcel->getActiveSheet()->getDefaultColumnDimension()->setWidth(8);
$objPHPExcel->getActiveSheet()->setShowGridlines(false);
		

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
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$8:$G8', NULL, 5),	//	Q1 to Q4
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
	$dataseriesLabels1,							    // plotLabel
	$xAxisTickValues1,								// plotCategory
	$dataSeriesValues1								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series1->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COLUMN);

$layout1 = new PHPExcel_Chart_Layout();
$layout1->getShowLeaderLines(); 

//	Set the series in the plot area
$plotarea1 = new PHPExcel_Chart_PlotArea($layout1, array($series1));

//	Set the series in the plot area
$plotarea1 = new PHPExcel_Chart_PlotArea(NULL, array($series1));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend1=NULL;
$title1 = new PHPExcel_Chart_Title('');
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
$chart1->setBottomRightPosition('O12');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart1);

$styleArray1 = array(
      'borders' => array(
        'bottom' => array(
          'style' => PHPExcel_Style_Border::BORDER_THIN,
          'color' => array('argb' => '6495ED'),
        ), 
      ),
    );

$styleArray2 = array(
      'borders' => array(
        'right' => array(
          'style' => PHPExcel_Style_Border::BORDER_THIN,
          'color' => array('argb' => '6495ED'),
        ), 
      ),
    );
	
$styleArray3 = array(
      'borders' => array(
        'top' => array(
          'style' => PHPExcel_Style_Border::BORDER_THIN,
          'color' => array('argb' => '6495ED'),
        ), 
      ),
    );

$styleArray4 = array(
      'borders' => array(
        'left' => array(
          'style' => PHPExcel_Style_Border::BORDER_THIN,
          'color' => array('argb' => '6495ED'),
        ), 
      ),
    );
	
	$objPHPExcel->getActiveSheet()->getStyle('C8:G8')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C8:G8')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C8:G8')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C8:G8')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C9:G9')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C9:G9')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C9:G9')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C9:G9')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C13:R13')->applyFromArray($styleArray1);

/////////////////////////////////////////////
//////////////////////////////////////////////
//////////////////////////////////////////////////
////////////////////////////////////////////////

$objWorksheet->fromArray(
$results3,NULL,'C19'
);

$dataseriesLabels2 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$19:$G19', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$objPHPExcel->getActiveSheet()->setCellValue('C18', 'Gross Profit');

$objPHPExcel->getActiveSheet()->getStyle("C18")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues2 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$19:$G19', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues2 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$20:$G20', NULL, 5));

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
$title2 = new PHPExcel_Chart_Title('');
$yAxisLabel2= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P20:P21');
$objPHPExcel->getActiveSheet()->mergeCells('R20:R21');
$objPHPExcel->getActiveSheet()->setCellValue('P18', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R18', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P19', '=IF(G20>C20,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R19', '=IF(G20>C20,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P20', '=round((G20-C20)/ABS(C20),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R20', '=round((G20-F20)/ABS(C20),2)');
$objPHPExcel->getActiveSheet()->getStyle('P19')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R19')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P20')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P20')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R20')->setConditionalStyles($conditionalStyles);
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

$chart2->setTopLeftPosition('I15');
$chart2->setBottomRightPosition('O23');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart2);

	$objPHPExcel->getActiveSheet()->getStyle('C19:G19')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C19:G19')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C19:G19')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C19:G19')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C20:G20')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C20:G20')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C20:G20')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C20:G20')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C24:R24')->applyFromArray($styleArray1);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////

$objWorksheet->fromArray(
$results27,NULL,'C30'
);

$dataseriesLabels3 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$30:$G30', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$objPHPExcel->getActiveSheet()->setCellValue('C29', 'Gross Profit %');
$objPHPExcel->getActiveSheet()->getStyle("C29")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');


$xAxisTickValues3 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$30:$G30', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues3 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$31:$G31', NULL, 5));

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
$title3 = new PHPExcel_Chart_Title('');
$yAxisLabel3= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P31:P32');
$objPHPExcel->getActiveSheet()->mergeCells('R31:R32');
$objPHPExcel->getActiveSheet()->setCellValue('P29', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R29', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P30', '=IF(G31>C31,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P31', '=round((G31-C31)/ABS(C31),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R31', '=round((G31-F31)/ABS(F31),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R30', '=IF(G31>C31,"↑","↓")');
$objPHPExcel->getActiveSheet()->getStyle('P30')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R30')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P31')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P31')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R31')->setConditionalStyles($conditionalStyles);
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

$chart3->setTopLeftPosition('I26');
$chart3->setBottomRightPosition('O34');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart3);
	$objPHPExcel->getActiveSheet()->getStyle('C30:G30')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C30:G30')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C30:G30')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C30:G30')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C31:G31')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C31:G31')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C31:G31')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C31:G31')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C35:R35')->applyFromArray($styleArray1);


/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////

$objWorksheet->fromArray(
$results4,NULL,'C41'
);

$dataseriesLabels4 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$41:$G41', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$objPHPExcel->getActiveSheet()->setCellValue('C40', 'Operating Income (ETBT)');

$objPHPExcel->getActiveSheet()->getStyle("C40")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');
$xAxisTickValues4 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$41:$G41', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues4 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$42:$G42', NULL, 5));

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
$title4 = new PHPExcel_Chart_Title('');

$yAxisLabel4= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P42:P43');
$objPHPExcel->getActiveSheet()->mergeCells('R42:R43');
$objPHPExcel->getActiveSheet()->setCellValue('P40', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R40', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P41', '=IF(G42>C42,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R41', '=IF(G42>F42,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P42', '=round((G42-C42)/ABS(C42),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R42', '=round((G42-F42)/ABS(F42),2)');
$objPHPExcel->getActiveSheet()->getStyle('P41')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R41')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P42')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P42')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R42')->setConditionalStyles($conditionalStyles);
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

$chart4->setTopLeftPosition('I37');
$chart4->setBottomRightPosition('O45');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart4);

	$objPHPExcel->getActiveSheet()->getStyle('C41:G41')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C41:G41')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C41:G41')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C41:G41')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C42:G42')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C42:G42')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C42:G42')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C42:G42')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C46:R46')->applyFromArray($styleArray1);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////


$objWorksheet->fromArray(
$results28, NULL, 'C52'
);

$dataseriesLabels5 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$52:$G52', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C51', 'Operating Income (ETBT) %');
$objPHPExcel->getActiveSheet()->getStyle("C51")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues5 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$52:$G52', NULL, 5),//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues5 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$53:$G53', NULL, 5));

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

//	Set the series in the plot area
$plotarea5 = new PHPExcel_Chart_PlotArea(NULL, array($series5));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend5=NULL;
$title5 = new PHPExcel_Chart_Title('');
$yAxisLabel5= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P53:P54');
$objPHPExcel->getActiveSheet()->mergeCells('R53:R54');
$objPHPExcel->getActiveSheet()->setCellValue('P51', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R51', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P52', '=IF(G53>C53,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R52', '=IF(G53>F53,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P53', '=round((G53-C53)/ABS(C53),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R53', '=round((G53-F53)/ABS(F53),2)');
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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P53')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P53')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R53')->setConditionalStyles($conditionalStyles);
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

$chart5->setTopLeftPosition('I48');
$chart5->setBottomRightPosition('O56');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart5);

	$objPHPExcel->getActiveSheet()->getStyle('C52:G52')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C52:G52')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C52:G52')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C52:G52')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C53:G53')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C53:G53')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C53:G53')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C53:G53')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C57:R57')->applyFromArray($styleArray1);


	
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////

$objWorksheet->fromArray(
$results6, NULL, 'C63'
);

$dataseriesLabels6 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$63:$G63', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C62', 'Income before tax');
$objPHPExcel->getActiveSheet()->getStyle("C62")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues6 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$63:$G63', NULL, 5),//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues6 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$64:$G64', NULL, 5));

//	Build the dataseries
$series6 = new PHPExcel_Chart_DataSeries(
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

//	Set the series in the plot area
$plotarea6 = new PHPExcel_Chart_PlotArea(NULL, array($series6));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend6=NULL;
$title6 = new PHPExcel_Chart_Title('');
$yAxisLabel6= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P65:P66');
$objPHPExcel->getActiveSheet()->mergeCells('R65:R66');
$objPHPExcel->getActiveSheet()->setCellValue('P63', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R63', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P64', '=IF(G64>C64,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R64', '=IF(G64>F64,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P65', '=round((G64-C64)/ABS(C64),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R65', '=round((G64-F64)/ABS(F64),2)');
$objPHPExcel->getActiveSheet()->getStyle('P65')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R65')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P65')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P65')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R65')->setConditionalStyles($conditionalStyles);
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

//	Set the position where the chart should appear in the Financial_Ratios

$chart6->setTopLeftPosition('I59');
$chart6->setBottomRightPosition('O68');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart6);


	$objPHPExcel->getActiveSheet()->getStyle('C63:G63')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C63:G63')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C63:G63')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C63:G63')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C64:G64')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C64:G64')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C64:G64')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C64:G64')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C69:R69')->applyFromArray($styleArray1);

/////////////////////
///////////////////////
//////////////////////////
///////////////////////////


$objWorksheet->fromArray(
$results29, NULL, 'C75'
);

$dataseriesLabels7 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$75:$G75', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C74', 'Income before tax %');
$objPHPExcel->getActiveSheet()->getStyle("C74")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues7 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$75:$G75', NULL, 5),//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues7 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$76:$G76', NULL, 5));

//	Build the dataseries
$series7 = new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues7)-1),			// plotOrder
	$dataseriesLabels7,							// plotLabel
	$xAxisTickValues7,								// 
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
$title7 = new PHPExcel_Chart_Title('');
$yAxisLabel7= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P77:P78');
$objPHPExcel->getActiveSheet()->mergeCells('R77:R78');
$objPHPExcel->getActiveSheet()->setCellValue('P75', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R75', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P76', '=IF(G76>C76,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R76', '=IF(G76>F76,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P77', '=round((G76-C76)/ABS(C76),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R77', '=round((G76-F76)/ABS(F76),2)');
$objPHPExcel->getActiveSheet()->getStyle('P76')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R76')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P77')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P77')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R77')->setConditionalStyles($conditionalStyles);
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

//	Set the position where the chart should appear in the Financial_Ratios

$chart7->setTopLeftPosition('I71');
$chart7->setBottomRightPosition('O79');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart7);


	$objPHPExcel->getActiveSheet()->getStyle('C75:G75')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C75:G75')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C75:G75')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C75:G75')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C76:G76')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C76:G76')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C76:G76')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C76:G76')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C80:R80')->applyFromArray($styleArray1);


/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////


$objWorksheet->fromArray(
$results7,NULL,'C86'
);

$dataseriesLabels8 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$86:$G86', NULL, 1),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$objPHPExcel->getActiveSheet()->setCellValue('C85', 'Net Profit');
$objPHPExcel->getActiveSheet()->getStyle("C85")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues8 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$86:$G86', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues8 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$87:$G87', NULL, 5));

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
$title8 = new PHPExcel_Chart_Title('');
$yAxisLabel8= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P87:P88');
$objPHPExcel->getActiveSheet()->mergeCells('R87:R88');
$objPHPExcel->getActiveSheet()->setCellValue('P85', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R85', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P86', '=IF(G87>C87,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R86', '=IF(G87>F87,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P87', '=round((G87-C87)/ABS(C87),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R87', '=round((G87-F87)/ABS(F87),2)');
$objPHPExcel->getActiveSheet()->getStyle('P86')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R86')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P87')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P87')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R87')->setConditionalStyles($conditionalStyles);
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

$chart8->setTopLeftPosition('I82');
$chart8->setBottomRightPosition('O90');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart8);

	$objPHPExcel->getActiveSheet()->getStyle('C86:G86')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C86:G86')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C86:G86')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C86:G86')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C87:G87')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C87:G87')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C87:G87')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C87:G87')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C91:R91')->applyFromArray($styleArray1);


/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////


$objWorksheet->fromArray(
$results30, NULL, 'C98'
);

$dataseriesLabels9 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$98:$G98', NULL, 1),	//	2010
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

$xAxisTickValues9 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$98:$G98', NULL, 5),//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues9 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$99:$G99', NULL, 5));

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

//	Set the series in the plot area
$plotarea9 = new PHPExcel_Chart_PlotArea(NULL, array($series9));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend9=NULL;
$title9 = new PHPExcel_Chart_Title('');
$yAxisLabel9= NULL;

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

//	Set the position where the chart should appear in the Financial_Ratios

$chart9->setTopLeftPosition('I93');
$chart9->setBottomRightPosition('O101');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart9);

	$objPHPExcel->getActiveSheet()->getStyle('C98:G98')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C98:G98')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C98:G98')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C98:G98')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C99:G99')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C99:G99')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C99:G99')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C99:G99')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C102:R102')->applyFromArray($styleArray1);


	/////////////////
	//////////////////
	/////////////////
	/////////////////
	/////////////////
	

$objWorksheet->fromArray(
$results8,NULL,'C108'
);

$dataseriesLabels10 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$108:$G108', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$objPHPExcel->getActiveSheet()->setCellValue('C107', 'Cash &  Equivalents');
$objPHPExcel->getActiveSheet()->getStyle("C107")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');


$xAxisTickValues10 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$108:$G108', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues10 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$109:$G109', NULL, 5));

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
$title10 = new PHPExcel_Chart_Title('');
$yAxisLabel10= NULL;


$objPHPExcel->getActiveSheet()->mergeCells('P108:P109');
$objPHPExcel->getActiveSheet()->mergeCells('R108:R109');
$objPHPExcel->getActiveSheet()->setCellValue('P106', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R106', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P107', '=IF(G109>C109,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R107', '=IF(G109>F109,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P108', '=round((G109-C109)/ABS(C109),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R108', '=round((G109-F109)/ABS(F109),2)');
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

$chart10->setTopLeftPosition('I104');
$chart10->setBottomRightPosition('O112');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart10);

$objPHPExcel->getActiveSheet()->getStyle('C108:G108')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C108:G108')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C108:G108')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C108:G108')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C109:G109')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C109:G109')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C109:G109')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C109:G109')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C113:R113')->applyFromArray($styleArray1);

//////////////////
//////////////////
//////////////////
//////////////////
//////////////////


$objWorksheet->fromArray(
$results9,NULL,'C119'
);

$dataseriesLabels11 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$119:$G119', NULL, 5),	//	2010
	);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C118', 'Account Receivable');
$objPHPExcel->getActiveSheet()->getStyle("C118")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues11 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$119:$G119', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues11 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$120:$G120', NULL, 5));

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
$title11 = new PHPExcel_Chart_Title('');
$yAxisLabel11= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P120:P121');
$objPHPExcel->getActiveSheet()->mergeCells('R120:R121');
$objPHPExcel->getActiveSheet()->setCellValue('P118', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R118', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P119', '=IF(G120>C120,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R119', '=IF(G120>F120,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P120', '=round((G120-C120)/ABS(C120),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R120', '=round((G120-F120)/ABS(F120),2)');
$objPHPExcel->getActiveSheet()->getStyle('P119')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R119')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P120')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P120')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R120')->setConditionalStyles($conditionalStyles);
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
	$yAxisLabel1	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart11->setTopLeftPosition('I115');
$chart11->setBottomRightPosition('O123');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart11);

	$objPHPExcel->getActiveSheet()->getStyle('C119:G119')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C119:G119')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C119:G119')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C119:G119')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C120:G120')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C120:G120')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C120:G120')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C120:G120')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C124:R124')->applyFromArray($styleArray1);
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////


$objWorksheet->fromArray(
$results10,NULL,'C130'
);

$dataseriesLabels12 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$130:$G130', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C129', 'Inventory');
$objPHPExcel->getActiveSheet()->getStyle("C129")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues12 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$130:$G130', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues12 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$131:$G131', NULL, 5));

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
$title12 = new PHPExcel_Chart_Title('');
$yAxisLabel12= NULL;


$objPHPExcel->getActiveSheet()->mergeCells('P131:P132');
$objPHPExcel->getActiveSheet()->mergeCells('R131:R132');
$objPHPExcel->getActiveSheet()->setCellValue('P129', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R129', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P130', '=IF(G131>C131,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R130', '=IF(G131>F131,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P131', '=round((G131-C131)/ABS(C131),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R131', '=round((G131-F131)/ABS(F131),2)');
$objPHPExcel->getActiveSheet()->getStyle('P130')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R130')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P131')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P131')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R131')->setConditionalStyles($conditionalStyles);
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

$chart12->setTopLeftPosition('I126');
$chart12->setBottomRightPosition('O134');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart12);

	$objPHPExcel->getActiveSheet()->getStyle('C130:G130')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C130:G130')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C130:G130')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C130:G130')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C131:G131')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C131:G131')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C131:G131')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C131:G131')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C135:R135')->applyFromArray($styleArray1);


/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////

$objWorksheet->fromArray(
$results13,NULL,'C141'
);

$dataseriesLabels13 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$141:$G141', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C140', 'Long Term Investments');
$objPHPExcel->getActiveSheet()->getStyle("C140")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues13 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$141:$G141', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues13 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$142:$G142', NULL, 5));

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
$title13 = new PHPExcel_Chart_Title('');
$yAxisLabel13= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P142:P143');
$objPHPExcel->getActiveSheet()->mergeCells('R142:R143');
$objPHPExcel->getActiveSheet()->setCellValue('P140', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R140', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P141', '=IF(G142>C142,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R141', '=IF(G142>F142,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P142', '=round((G142-C142)/ABS(C142),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R142', '=round((G142-F142)/ABS(F142),2)');
$objPHPExcel->getActiveSheet()->getStyle('P141')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R141')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P142')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P142')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R142')->setConditionalStyles($conditionalStyles);
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

$chart13->setTopLeftPosition('I137');
$chart13->setBottomRightPosition('O145');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart13);
	$objPHPExcel->getActiveSheet()->getStyle('C141:G141')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C141:G141')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C141:G141')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C141:G141')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C142:G142')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C142:G142')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C142:G142')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C142:G142')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C146:R146')->applyFromArray($styleArray1);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////


$objWorksheet->fromArray(
$results15,NULL,'C152'
);

$dataseriesLabels15 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$152:$G152', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C151', 'Total Assets');
$objPHPExcel->getActiveSheet()->getStyle("C151")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues15 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$152:$G152', NULL, 5),//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues15 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$153:$G153', NULL, 5));

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
$plotarea15 = new PHPExcel_Chart_PlotArea(NULL, array($series15));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend15=NULL;
$title15 = new PHPExcel_Chart_Title('');
$yAxisLabel15= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P154:P155');
$objPHPExcel->getActiveSheet()->mergeCells('R154:R155');
$objPHPExcel->getActiveSheet()->setCellValue('P152', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R152', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P153', '=IF(G153>C153,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R153', '=IF(G153>F153,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P154', '=round((G153-C153)/ABS(C153),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R154', '=round((G153-F153)/ABS(F153),2)');
$objPHPExcel->getActiveSheet()->getStyle('P153')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R153')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P154')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P154')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R154')->setConditionalStyles($conditionalStyles);
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

$chart15->setTopLeftPosition('I148');
$chart15->setBottomRightPosition('O156');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart15);

	$objPHPExcel->getActiveSheet()->getStyle('C152:G152')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C152:G152')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C152:G152')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C152:G152')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C153:G153')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C153:G153')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C153:G153')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C153:G153')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C157:R157')->applyFromArray($styleArray1);

//////////////
///////////////
////////////////
///////////////

$objWorksheet->fromArray(
$results16,NULL,'C163'
);

$dataseriesLabels16 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$163:$G163', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C162', 'Accounts Payable');
$objPHPExcel->getActiveSheet()->getStyle("C162")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues16 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$163:$G163', NULL, 5),//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues16 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$164:$G164', NULL, 5));

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
$title16 = new PHPExcel_Chart_Title('');
$yAxisLabel16= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P165:P166');
$objPHPExcel->getActiveSheet()->mergeCells('R165:R166');
$objPHPExcel->getActiveSheet()->setCellValue('P163', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R163', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P164', '=IF(G164>C164,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R164', '=IF(G164>F164,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P165', '=round((G164-C164)/ABS(C164),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R165', '=round((G164-F164)/ABS(F164),2)');
$objPHPExcel->getActiveSheet()->getStyle('P164')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R164')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P165')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P165')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R165')->setConditionalStyles($conditionalStyles);
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

$chart16->setTopLeftPosition('I159');
$chart16->setBottomRightPosition('O167');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart16);

	$objPHPExcel->getActiveSheet()->getStyle('C163:G163')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C163:G163')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C163:G163')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C163:G163')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C164:G164')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C164:G164')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C164:G164')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C164:G164')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C168:R168')->applyFromArray($styleArray1);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////

$objWorksheet->fromArray(
$results17,NULL,'C174'
);

$dataseriesLabels17 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C174:$G174', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C173', 'Current Liabilities');
$objPHPExcel->getActiveSheet()->getStyle("C173")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues17 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$174:$G174', NULL, 5),//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values	
//		Data Marker
$dataSeriesValues17 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$175:$G175', NULL, 5));

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
$title17 = new PHPExcel_Chart_Title('');
$yAxisLabel17= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P175:P176');
$objPHPExcel->getActiveSheet()->mergeCells('R175:R176');
$objPHPExcel->getActiveSheet()->setCellValue('P173', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R173', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P174', '=IF(G175>C175,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R174', '=IF(G175>F175,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P175', '=round((G175-C175)/ABS(C175),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R175', '=round((G175-F175)/ABS(F175),2)');
$objPHPExcel->getActiveSheet()->getStyle('P174')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R174')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P175')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P175')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R175')->setConditionalStyles($conditionalStyles);
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

$chart17->setTopLeftPosition('I170');
$chart17->setBottomRightPosition('O178');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart17);

	$objPHPExcel->getActiveSheet()->getStyle('C174:G174')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C174:G174')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C174:G174')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C174:G174')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C175:G175')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C175:G175')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C175:G175')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C175:G175')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C179:R179')->applyFromArray($styleArray1);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////


$objWorksheet->fromArray(
$results18, NULL, 'C185'
);

$dataseriesLabels18 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$185:$G185', NULL, 1),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C184', 'Long Term Debt');
$objPHPExcel->getActiveSheet()->getStyle("C184")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues18 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$185:$G185', NULL, 5),//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues18 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$186:$G186', NULL, 5));

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
$title18 = new PHPExcel_Chart_Title('');
$yAxisLabel18= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P186:P187');
$objPHPExcel->getActiveSheet()->mergeCells('R186:R187');
$objPHPExcel->getActiveSheet()->setCellValue('P184', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R184', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P185', '=IF(G186>C186,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R185', '=IF(G186>F186,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P186', '=round((G186-C186)/ABS(C186),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R186', '=round((G186-F186)/ABS(F186),2)');
$objPHPExcel->getActiveSheet()->getStyle('P185')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R185')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P186')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P186')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R186')->setConditionalStyles($conditionalStyles);
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

$chart18->setTopLeftPosition('I181');
$chart18->setBottomRightPosition('O189');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart18);

	$objPHPExcel->getActiveSheet()->getStyle('C185:G185')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C185:G185')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C185:G185')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C185:G185')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C186:G186')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C186:G186')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C186:G186')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C186:G186')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C190:R190')->applyFromArray($styleArray1);


/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////

$objWorksheet->fromArray(
$results20, NULL, 'C196'
);

$dataseriesLabels19 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$196:$G196', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C195', 'Total Liabilities');
$objPHPExcel->getActiveSheet()->getStyle("C195")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues19 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$196:$G196', NULL, 5),//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues19 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$197:$G197', NULL, 5));

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
$title19 = new PHPExcel_Chart_Title('');
$yAxisLabel19= NULL;


$objPHPExcel->getActiveSheet()->mergeCells('P198:P199');
$objPHPExcel->getActiveSheet()->mergeCells('R198:R199');
$objPHPExcel->getActiveSheet()->setCellValue('P196', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R196', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P197', '=IF(G197>C197,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R197', '=IF(G197>F197,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P198', '=round((G197-C197)/ABS(C197),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R198', '=round((G197-F197)/ABS(F197),2)');
$objPHPExcel->getActiveSheet()->getStyle('P197')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R197')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P198')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P198')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R198')->setConditionalStyles($conditionalStyles);
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

$chart19->setTopLeftPosition('I192');
$chart19->setBottomRightPosition('O200');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart19);


	$objPHPExcel->getActiveSheet()->getStyle('C196:G196')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C196:G196')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C196:G196')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C196:G196')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C197:G197')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C197:G197')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C197:G197')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C197:G197')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C201:R201')->applyFromArray($styleArray1);

	
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////





$objWorksheet->fromArray(
$results21, NULL, 'C207'
);

$dataseriesLabels20 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$207:$G207', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C206', 'Total Equity');
$objPHPExcel->getActiveSheet()->getStyle("C206")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues20 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$207:$G207', NULL, 5),//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues20 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$207:$G207', NULL, 5));

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
$title20 = new PHPExcel_Chart_Title('Total Equity');
$yAxisLabel20= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P209:P210');
$objPHPExcel->getActiveSheet()->mergeCells('R209:R210');
$objPHPExcel->getActiveSheet()->setCellValue('P207', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R207', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P208', '=IF(G208>C208,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R208', '=IF(G208>F208,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P209', '=round((G208-C208)/ABS(C208),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R209', '=round((G208-F208)/ABS(F208),2)');
$objPHPExcel->getActiveSheet()->getStyle('P208')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R208')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P209')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P209')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R209')->setConditionalStyles($conditionalStyles);
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

$chart20->setTopLeftPosition('I203');
$chart20->setBottomRightPosition('O211');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart20);

	$objPHPExcel->getActiveSheet()->getStyle('C207:G207')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C207:G207')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C207:G207')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C207:G207')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C208:G208')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C208:G208')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C208:G208')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C208:G208')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C212:R212')->applyFromArray($styleArray1);


/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////


$objWorksheet->fromArray(
$results22, NULL, 'C217'
);

$dataseriesLabels21 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$217:$G217', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C216', 'Total Liabilities & Equity');
$objPHPExcel->getActiveSheet()->getStyle("C216")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues21 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$217:$G217', NULL, 5),//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues21 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$218:$G218', NULL, 5));

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
$title21 = new PHPExcel_Chart_Title('');

$yAxisLabel21= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P219:P220');
$objPHPExcel->getActiveSheet()->mergeCells('R218:R219');
$objPHPExcel->getActiveSheet()->setCellValue('P216', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R216', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P217', '=IF(G218>C218,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R217', '=IF(G218>F218,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P218', '=round((G218-C218)/ABS(C218),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R218', '=round((G218-F218)/ABS(F218),2)');
$objPHPExcel->getActiveSheet()->getStyle('P217')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R217')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$objPHPExcel->getActiveSheet()->getStyle('P218')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R218')->setConditionalStyles($conditionalStyles);
//conditions


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

$chart21->setTopLeftPosition('I213');
$chart21->setBottomRightPosition('O221');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart21);

	$objPHPExcel->getActiveSheet()->getStyle('C217:G217')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C217:G217')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C217:G217')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C217:G217')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C218:G218')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C218:G218')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C218:G218')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C218:G218')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C222:R222')->applyFromArray($styleArray1);




/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////


$objWorksheet->fromArray(
$results23, NULL, 'C227'
);

$dataseriesLabels22 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$227:$G227', NULL, 5),	//	2010
	);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C226', 'Cash from Operating Activities');
$objPHPExcel->getActiveSheet()->getStyle("C226")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues22 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$227:$G227', NULL, 5),//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues22 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$228:$G228', NULL, 5));

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
$plotarea22 = new PHPExcel_Chart_PlotArea(NULL, array($series22));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend22=NULL;
$title22 = new PHPExcel_Chart_Title('');
$yAxisLabel22= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P228:P229');
$objPHPExcel->getActiveSheet()->mergeCells('R228:R229');
$objPHPExcel->getActiveSheet()->setCellValue('P226', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R226', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P227', '=IF(G228>C228,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R227', '=IF(G228>F228,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P228', '=round((G228-C228)/ABS(C228),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R228', '=round((G227-F228)/ABS(F228),2)');
$objPHPExcel->getActiveSheet()->getStyle('P227')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R227')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P228')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P228')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R228')->setConditionalStyles($conditionalStyles);
//conditions

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

$chart22->setTopLeftPosition('I224');
$chart22->setBottomRightPosition('O232');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart22);

	$objPHPExcel->getActiveSheet()->getStyle('C227:G227')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C227:G227')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C227:G227')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C227:G227')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C228:G228')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C228:G228')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C228:G228')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C228:G228')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C233:R233')->applyFromArray($styleArray1);



/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////


$objWorksheet->fromArray(
$results24, NULL, 'C239'
);

$dataseriesLabels23 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$239:$G239', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C238', 'Cash from Investing Activities');
$objPHPExcel->getActiveSheet()->getStyle("C238")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues23 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$239:$G239', NULL, 5),//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues23 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$240:$G240', NULL, 5));

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
$title23 = new PHPExcel_Chart_Title('');
$yAxisLabel23= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P240:P241');
$objPHPExcel->getActiveSheet()->mergeCells('R240:R241');
$objPHPExcel->getActiveSheet()->setCellValue('P238', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R238', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P239', '=IF(G240>C240,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R239', '=IF(G240>F240,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P240', '=round((G240-C240)/ABS(C240),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R240', '=round((G240-F240)/ABS(F240),2)');
$objPHPExcel->getActiveSheet()->getStyle('P239')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R239')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P240')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P240')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R240')->setConditionalStyles($conditionalStyles);
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

$chart23->setTopLeftPosition('I236');
$chart23->setBottomRightPosition('O244');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart23);

	$objPHPExcel->getActiveSheet()->getStyle('C239:G239')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C239:G239')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C239:G239')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C239:G239')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C240:G240')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C240:G240')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C240:G240')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C240:G240')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C245:R245')->applyFromArray($styleArray1);



/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////


$objWorksheet->fromArray(
$results25, NULL, 'C251'
);

$dataseriesLabels24 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$251:$G251', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C250', 'Cash from Financing Activities');
$objPHPExcel->getActiveSheet()->getStyle("C250")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues24 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$252:$G252', NULL, 5),//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues24 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$252:$G252', NULL, 5));

//	Build the dataseries
$series24= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues24)-1),			// plotOrder
	$dataseriesLabels24,								// plotLabel
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
$title24 = new PHPExcel_Chart_Title('');
$yAxisLabel24= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P252:P253');
$objPHPExcel->getActiveSheet()->mergeCells('R252:R253');
$objPHPExcel->getActiveSheet()->setCellValue('P250', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R250', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P251', '=IF(G252>C252,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R251', '=IF(G252>F252,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P252', '=round((G252-C252)/ABS(C252),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R252', '=round((G252-F252)/ABS(F252),2)');
$objPHPExcel->getActiveSheet()->getStyle('P251')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R251')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P252')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P252')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R252')->setConditionalStyles($conditionalStyles);
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

$chart24->setTopLeftPosition('I247');
$chart24->setBottomRightPosition('O255');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart24);

	$objPHPExcel->getActiveSheet()->getStyle('C251:G251')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C251:G251')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C251:G251')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C251:G251')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C252:G252')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C252:G252')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C252:G252')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C252:G252')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C256:R256')->applyFromArray($styleArray1);




/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////

$objWorksheet->fromArray(
$results26, NULL, 'C262'
);

$dataseriesLabels25 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$262:$G262', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C261', 'Net Change in Cash');
$objPHPExcel->getActiveSheet()->getStyle("C261")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues25 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$262:$G262', NULL, 5),//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues25 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$263:$G263', NULL, 5));

//	Build the dataseries
$series25= new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues25)-1),			// plotOrder
	$dataseriesLabels25,								// plotLabel
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
$title25 = new PHPExcel_Chart_Title('');
$yAxisLabel25= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P263:P264');
$objPHPExcel->getActiveSheet()->mergeCells('R263:R264');
$objPHPExcel->getActiveSheet()->setCellValue('P261', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R261', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P262', '=IF(G263>C263,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R262', '=IF(G263>F263,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P263', '=round((G263-C263)/ABS(C263),2)');
$objPHPExcel->getActiveSheet()->setCellValue('R263', '=round((G263-F263)/ABS(F263),2)');
$objPHPExcel->getActiveSheet()->getStyle('P262')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R262')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P263')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P263')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R263')->setConditionalStyles($conditionalStyles);
//conditions


//	Create the chart
$chart25= new PHPExcel_Chart(
	'chart25',		// name
	$title25,		// title
	$legend25,		// legend
	$plotarea25,		// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	NULL,			// xAxisLabel
	$yAxisLabel25	// yAxisLabel
);

//	Set the position where the chart should appear in the Financial_Ratios

$chart25->setTopLeftPosition('I258');
$chart25->setBottomRightPosition('O266');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart25);

	$objPHPExcel->getActiveSheet()->getStyle('C262:G262')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C262:G262')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C262:G262')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C262:G262')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C263:G263')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C263:G263')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C263:G263')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C263:G263')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C267:R267')->applyFromArray($styleArray1);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
$objWorksheet->fromArray(
$results31, NULL, 'C273'
);

$dataseriesLabels26 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$273:$G273', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C272', 'Current Ratio');
$objPHPExcel->getActiveSheet()->getStyle("C272")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues26 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$273:$G273', NULL, 5),//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues26 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$274:$G274', NULL, 5));

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
$title26 = new PHPExcel_Chart_Title('');
$yAxisLabel26= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P275:P276');
$objPHPExcel->getActiveSheet()->mergeCells('R275:R276');
$objPHPExcel->getActiveSheet()->setCellValue('P273', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R273', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P274', '=IF(G274>C274,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R274', '=IF(G274>F274,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P275', '=(G274-C274)/ABS(C274)');
$objPHPExcel->getActiveSheet()->setCellValue('R275', '=(G274-F274)/ABS(F274)');
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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P275')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P275')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R275')->setConditionalStyles($conditionalStyles);
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

$chart26->setTopLeftPosition('I269');
$chart26->setBottomRightPosition('O277');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart26);

$objPHPExcel->getActiveSheet()->getStyle('C273:G273')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C273:G273')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C273:G273')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C273:G273')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C274:G274')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C274:G274')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C274:G274')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C274:G274')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C278:R278')->applyFromArray($styleArray1);



/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////



$objWorksheet->fromArray(
$results32, NULL, 'C284'
);

$dataseriesLabels27 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$284:$G284', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C283', 'Quick Ratio');
$objPHPExcel->getActiveSheet()->getStyle("C283")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues27 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$284:$G284', NULL, 5),//	Q1 to Q4
);
	
//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues27 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$285:$G285', NULL, 5));

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
$title27 = new PHPExcel_Chart_Title('');
$yAxisLabel27= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P285:P286');
$objPHPExcel->getActiveSheet()->mergeCells('R285:R286');
$objPHPExcel->getActiveSheet()->setCellValue('P283', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R283', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P284', '=IF(G285>C285,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R284', '=IF(G285>F285,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P285', '=(G285-C285)/ABS(C285)');
$objPHPExcel->getActiveSheet()->setCellValue('R285', '=(G285-F285)/ABS(F285)');
$objPHPExcel->getActiveSheet()->getStyle('P284')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R284')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P285')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P285')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R285')->setConditionalStyles($conditionalStyles);
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

$chart27->setTopLeftPosition('I280');
$chart27->setBottomRightPosition('O289');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart27);

	$objPHPExcel->getActiveSheet()->getStyle('C284:G284')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C284:G284')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C284:G284')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C284:G284')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C285:G285')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C285:G285')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C285:G285')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C285:G285')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C290:R290')->applyFromArray($styleArray1);


	///////////////
	///////////////
	//////////////
	///////////////


$objWorksheet->fromArray(
$results33, NULL, 'C296'
);

$dataseriesLabels28 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$296:$G296', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C295', 'Assets Turnover');
$objPHPExcel->getActiveSheet()->getStyle("C295")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues28 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$296:$G296', NULL, 5),//	Q1 to Q4
);
	
//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues28 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$297:$G297', NULL, 5));

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
$title28 = new PHPExcel_Chart_Title('');
$yAxisLabel28= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P297:P298');
$objPHPExcel->getActiveSheet()->mergeCells('R297:R298');
$objPHPExcel->getActiveSheet()->setCellValue('P295', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R295', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P296', '=IF(G297>C297,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R296', '=IF(G297>F297,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P297', '=(G297-C297)/ABS(C297)');
$objPHPExcel->getActiveSheet()->setCellValue('R297', '=(G297-F297)/ABS(F297)');
$objPHPExcel->getActiveSheet()->getStyle('P297')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R297')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P297')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P297')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R297')->setConditionalStyles($conditionalStyles);
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

$chart28->setTopLeftPosition('I292');
$chart28->setBottomRightPosition('O300');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart28);

	$objPHPExcel->getActiveSheet()->getStyle('C296:G296')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C296:G296')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C296:G296')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C296:G296')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C297:G297')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C297:G297')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C297:G297')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C297:G297')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C301:R301')->applyFromArray($styleArray1);


	///////////////
	///////////////
	//////////////
	///////////////


$objWorksheet->fromArray(
$results34, NULL, 'C307'
);

$dataseriesLabels29 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$307:$G307', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C306', 'Inventory Turnover');
$objPHPExcel->getActiveSheet()->getStyle("C306")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues29 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$307:$G307', NULL, 5),//	Q1 to Q4
);
	
//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues29 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$308:$G308', NULL, 5));

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
$title29 = new PHPExcel_Chart_Title('');
$yAxisLabel29= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P308:P309');
$objPHPExcel->getActiveSheet()->mergeCells('R308:R309');
$objPHPExcel->getActiveSheet()->setCellValue('P306', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R306', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P307', '=IF(G308>C308,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R307', '=IF(G308>F308,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P308', '=(G308-C308)/ABS(C308)');
$objPHPExcel->getActiveSheet()->setCellValue('R308', '=(G308-F308)/ABS(F308)');
$objPHPExcel->getActiveSheet()->getStyle('P307')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R307')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P308')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P308')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R308')->setConditionalStyles($conditionalStyles);
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

$chart29->setTopLeftPosition('I303');
$chart29->setBottomRightPosition('O311');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart29);

	$objPHPExcel->getActiveSheet()->getStyle('C307:G307')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C307:G307')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C307:G307')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C307:G307')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C308:G308')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C308:G308')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C308:G308')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C308:G308')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C312:R312')->applyFromArray($styleArray1);


	///////////////
	///////////////
	//////////////
	///////////////

	
	


$objWorksheet->fromArray(
$results35, NULL, 'C318'
);

$dataseriesLabels30 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$318:$G318', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C317', 'Accounts Receivable Turnover');
$objPHPExcel->getActiveSheet()->getStyle("C317")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues30 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$318:$G318', NULL, 5),//	Q1 to Q4
);
	
//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues30 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$319:$G319', NULL, 5));

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
$title30 = new PHPExcel_Chart_Title('');
$yAxisLabel30= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P319:P320');
$objPHPExcel->getActiveSheet()->mergeCells('R319:R320');
$objPHPExcel->getActiveSheet()->setCellValue('P317', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R317', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P318', '=IF(G319>C319,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R318', '=IF(G319>F319,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P319', '=(G319-C319)/ABS(C319)');
$objPHPExcel->getActiveSheet()->setCellValue('R319', '=(G319-F319)/ABS(F319)');
$objPHPExcel->getActiveSheet()->getStyle('P318')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R318')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P319')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P319')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R319')->setConditionalStyles($conditionalStyles);
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

$chart30->setTopLeftPosition('I314');
$chart30->setBottomRightPosition('O322');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart30);

	$objPHPExcel->getActiveSheet()->getStyle('C318:G318')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C318:G318')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C318:G318')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C318:G318')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C319:G319')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C319:G319')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C319:G319')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C319:G319')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C323:R323')->applyFromArray($styleArray1);



/////////
/////////
/////////
//////////
	


$objWorksheet->fromArray(
$results36, NULL, 'C329'
);

$dataseriesLabels31 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$329:$G329', NULL, 5),	//	2010
	
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C328', 'Average Collection Period (Days)');
$objPHPExcel->getActiveSheet()->getStyle("C328")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues31 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$329:$G329', NULL, 5),//	Q1 to Q4
);
	
//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues31 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$330:$G330', NULL, 5));

//	Build the dataseries
$series31= new PHPExcel_Chart_DataSeries(
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
$title31 = new PHPExcel_Chart_Title('');
$yAxisLabel31= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P330:P331');
$objPHPExcel->getActiveSheet()->mergeCells('R330:R331');
$objPHPExcel->getActiveSheet()->setCellValue('P328', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R328', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P329', '=IF(G330>C330,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R329', '=IF(G330>F330,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P330', '=(G330-C330)/ABS(C330)');
$objPHPExcel->getActiveSheet()->setCellValue('R330', '=(G330-F330)/ABS(F330)');
$objPHPExcel->getActiveSheet()->getStyle('P329')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R329')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P330')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P330')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R330')->setConditionalStyles($conditionalStyles);
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

$chart31->setTopLeftPosition('I325');
$chart31->setBottomRightPosition('O333');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart31);

	$objPHPExcel->getActiveSheet()->getStyle('C329:G329')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C329:G329')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C329:G329')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C329:G329')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C330:G330')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C330:G330')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C330:G330')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C330:G330')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C334:R334')->applyFromArray($styleArray1);


/////////
/////////
/////////
//////////


$objWorksheet->fromArray(
$results37, NULL, 'C340'
);

$dataseriesLabels32 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$340:$G340', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C339', 'Inventory Period (Days)');
$objPHPExcel->getActiveSheet()->getStyle("C339")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues32 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$340:$G340', NULL, 5),//	Q1 to Q4
);
	
//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues32 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$341:$G341', NULL, 5));

//	Build the dataseries
$series32 = new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,	// plotGrouping
	range(0, count($dataSeriesValues32)-1),			// plotOrder
	$dataseriesLabels32,							// plotLabel
	$xAxisTickValues32,								// plotCategory
	$dataSeriesValues32								// plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series32->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

//	Set the series in the plot area
$plotarea32 = new PHPExcel_Chart_PlotArea(NULL, array($series32));
//	Set the chart legend
//$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$legend32=NULL;
$title32 = new PHPExcel_Chart_Title('');
$yAxisLabel32= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P342:P343');
$objPHPExcel->getActiveSheet()->mergeCells('R342:R343');
$objPHPExcel->getActiveSheet()->setCellValue('P340', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R340', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P341', '=IF(G341>C341,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R341', '=IF(G341>F341,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P342', '=(G341-C341)/ABS(C341)');
$objPHPExcel->getActiveSheet()->setCellValue('R342', '=(G341-F341)/ABS(F341)');
$objPHPExcel->getActiveSheet()->getStyle('P341')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R341')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P342')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P342')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R342')->setConditionalStyles($conditionalStyles);
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

$chart32->setTopLeftPosition('I336');
$chart32->setBottomRightPosition('O344');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart32);

	$objPHPExcel->getActiveSheet()->getStyle('C340:G340')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C340:G340')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C340:G340')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C340:G340')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C341:G341')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C341:G341')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C341:G341')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C341:G341')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C345:R345')->applyFromArray($styleArray1);




/////////
/////////
/////////
//////////


$objWorksheet->fromArray(
$results38, NULL, 'C351'
);

$dataseriesLabels33 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$351:$G351', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C350', 'Debt Ratio');
$objPHPExcel->getActiveSheet()->getStyle("C350")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues33 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$351:$G351', NULL, 5),//	Q1 to Q4
);
	
//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues33 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$352:$G352', NULL, 5));

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
$title33 = new PHPExcel_Chart_Title('');
$yAxisLabel33= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P352:P353');
$objPHPExcel->getActiveSheet()->mergeCells('R352:R353');
$objPHPExcel->getActiveSheet()->setCellValue('P350', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R350', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P351', '=IF(G352>C352,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R351', '=IF(G352>F352,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P352', '=(G352-C352)/ABS(C352)');
$objPHPExcel->getActiveSheet()->setCellValue('R352', '=(G352-F352)/ABS(F352)');
$objPHPExcel->getActiveSheet()->getStyle('P351')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R351')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P352')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P352')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R352')->setConditionalStyles($conditionalStyles);
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

$chart33->setTopLeftPosition('I347');
$chart33->setBottomRightPosition('O355');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart33);

	$objPHPExcel->getActiveSheet()->getStyle('C351:G351')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C351:G351')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C351:G351')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C351:G351')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C352:G352')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C352:G352')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C352:G352')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C352:G352')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C356:R356')->applyFromArray($styleArray1);


/////////
/////////
/////////
//////////
/////////
/////////
/////////
//////////

$objWorksheet->fromArray(
$results39, NULL, 'C362'
);

$dataseriesLabels34 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$358:$G358', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C361', 'Debt To Equity Ratio');
$objPHPExcel->getActiveSheet()->getStyle("C361")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues34 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$362:$G362', NULL, 5),//	Q1 to Q4
);
	
//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues34 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$363:$G363', NULL, 5));

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
$title34 = new PHPExcel_Chart_Title('');
$yAxisLabel34= NULL;

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

$chart34->setTopLeftPosition('I358');
$chart34->setBottomRightPosition('O366');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart34);

	$objPHPExcel->getActiveSheet()->getStyle('C362:G362')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C362:G362')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C362:G362')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C362:G362')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C363:G363')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C363:G363')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C363:G363')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C363:G363')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C367:R367')->applyFromArray($styleArray1);

/////////
/////////
/////////
//////////

$objWorksheet->fromArray(
$results40, NULL, 'C373'
);

$dataseriesLabels35 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$373:$G373', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C372', 'Return on Assets % ');
$objPHPExcel->getActiveSheet()->getStyle("C372")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues35 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$373:$G373', NULL, 5),//	Q1 to Q4
);
	
//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues35 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$374:$G374', NULL, 5));

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
$title35 = new PHPExcel_Chart_Title('');
$yAxisLabel35= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P373:P374');
$objPHPExcel->getActiveSheet()->mergeCells('R373:R374');
$objPHPExcel->getActiveSheet()->setCellValue('P371', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R371', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P372', '=IF(G374>C374,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R372', '=IF(G374>F374,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P373', '=(G374-C374)/ABS(C374)');
$objPHPExcel->getActiveSheet()->setCellValue('R373', '=(G374-F374)/ABS(F374)');
$objPHPExcel->getActiveSheet()->getStyle('P372')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R372')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P373')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P373')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R373')->setConditionalStyles($conditionalStyles);
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

$chart35->setTopLeftPosition('I369');
$chart35->setBottomRightPosition('O377');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart35);

	$objPHPExcel->getActiveSheet()->getStyle('C373:G373')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C373:G373')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C373:G373')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C373:G373')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C374:G374')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C374:G374')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C374:G374')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C374:G374')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C378:R378')->applyFromArray($styleArray1);
/////////
/////////
/////////
//////////



$objWorksheet->fromArray(
$results42, NULL, 'C384'
);

$dataseriesLabels36 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$384:$G384', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C383', 'Return on Equity %');
$objPHPExcel->getActiveSheet()->getStyle("C383")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues36 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$384:$G384', NULL, 5),//	Q1 to Q4
);
	
//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues36 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$385:$G385', NULL, 5));

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
$title36 = new PHPExcel_Chart_Title('');
$yAxisLabel36= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P384:P385');
$objPHPExcel->getActiveSheet()->mergeCells('R384:R385');
$objPHPExcel->getActiveSheet()->setCellValue('P382', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R382', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P383', '=IF(G385>C385,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R383', '=IF(G385>F385,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P384', '=(G385-C385)/ABS(C385)');
$objPHPExcel->getActiveSheet()->setCellValue('R384', '=(G385-F385)/ABS(F385)');
$objPHPExcel->getActiveSheet()->getStyle('P383')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R383')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P384')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P384')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R384')->setConditionalStyles($conditionalStyles);
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

$chart36->setTopLeftPosition('I380');
$chart36->setBottomRightPosition('O388');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart36);

	$objPHPExcel->getActiveSheet()->getStyle('C384:G384')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C384:G384')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C384:G384')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C384:G384')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C385:G385')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C385:G385')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C385:G385')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C385:G385')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C389:R389')->applyFromArray($styleArray1);
/////////
//////////
/////////
/////////
/////////
//////////


$objWorksheet->fromArray(
$results41, NULL, 'C395'
);

$dataseriesLabels37 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$395:$G395', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$objPHPExcel->getActiveSheet()->setCellValue('C394', 'Net Working Capital Ratio');
$objPHPExcel->getActiveSheet()->getStyle("C394")->getFont()->setBold(false)
                                ->setName('Helvetica')
                                ->setSize(14)
								->getColor()->setRGB('6a5acd');

$xAxisTickValues37 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Ratios!$C$395:$G395', NULL, 5),//	Q1 to Q4
);
	
//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues37 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Ratios!$C$396:$G396', NULL, 5));

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
$title37 = new PHPExcel_Chart_Title('');
$yAxisLabel37= NULL;

$objPHPExcel->getActiveSheet()->mergeCells('P396:P397');
$objPHPExcel->getActiveSheet()->mergeCells('R396:R397');
$objPHPExcel->getActiveSheet()->setCellValue('P394', '5 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('R394', '1 Yr. Trend');
$objPHPExcel->getActiveSheet()->setCellValue('P395', '=IF(G396>C396,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('R395', '=IF(G396>F396,"↑","↓")');
$objPHPExcel->getActiveSheet()->setCellValue('P396', '=(G396-C396)/ABS(C396)');
$objPHPExcel->getActiveSheet()->setCellValue('R396', '=(G396-F396)/ABS(F396)');
$objPHPExcel->getActiveSheet()->getStyle('P395')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('R395')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

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
$conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle('P396')->getConditionalStyles();
array_push($conditionalStyles, $objConditional1);
array_push($conditionalStyles, $objConditional3);
$objPHPExcel->getActiveSheet()->getStyle('P396')->setConditionalStyles($conditionalStyles);
$objPHPExcel->getActiveSheet()->getStyle('R396')->setConditionalStyles($conditionalStyles);
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

$chart37->setTopLeftPosition('I391');
$chart37->setBottomRightPosition('O399');

//	Add the chart to the Financial_Ratios
$objWorksheet->addChart($chart37);

	$objPHPExcel->getActiveSheet()->getStyle('C395:G395')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C395:G395')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C395:G395')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C395:G395')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C396:G396')->applyFromArray($styleArray1);
	$objPHPExcel->getActiveSheet()->getStyle('C396:G396')->applyFromArray($styleArray2);
	$objPHPExcel->getActiveSheet()->getStyle('C396:G396')->applyFromArray($styleArray3);
	$objPHPExcel->getActiveSheet()->getStyle('C396:G396')->applyFromArray($styleArray4);
	$objPHPExcel->getActiveSheet()->getStyle('C400:R400')->applyFromArray($styleArray1);


//$objPHPExcel->getActiveSheet()->setTitle('Financial_Ratios');
/////////
/////////
/////////
//////////

$newSheet=new PHPExcel_Worksheet($objPHPExcel,'Du_Pont_Analysis');
$objPHPExcel->addSheet($newSheet,1);
$objWorksheet=$objPHPExcel->setActiveSheetIndex(1);  
$objPHPExcel->getActiveSheet()->getDefaultColumnDimension()->setWidth(8);
$objPHPExcel->getActiveSheet()->setShowGridlines(false);

$query1 = "select income_statement.finyear, income_statement.totalrevenue_sales,income_statement.costrevenue,
income_statement.netprofit, 
balance_sheet.totalassets, balance_sheet.inventory, balance_sheet.accountrecive, balance_sheet.totalequity from income_statement, 
balance_sheet where income_statement.company='".$companyname."' and income_statement.finyear=balance_sheet.finyear 
and income_statement.company=balance_sheet.company";

$resultnew = mysqli_query($conn,$query1);
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
while ($info5 = mysqli_fetch_array($resultnew,MYSQLI_ASSOC))
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
$objPHPExcel->getActiveSheet()->setCellValue('J7', '=round((M9-I9)/ABS(I9),2)');
$objPHPExcel->getActiveSheet()->setCellValue('L7', '=round((M9-L9)/ABS(L9),2)');
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

$chart1->setTopLeftPosition('J10');
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
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$D$23:$H23', NULL, 5),	//	2010
	
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
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$D$23:$H23', NULL, 5),	//	Q1 to Q4
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

$chart2->setTopLeftPosition('E25');
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
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$N$23:$R23', NULL, 5),	//	2010

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
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$N$23:$R23', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues3 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Du_Pont_Analysis!$N$24:$R24', NULL, 5),
	);

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

$chart3->setTopLeftPosition('O25');
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
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$38:$E38', NULL, 5),	//	2010
	
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
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$38:$E38', NULL, 5),	//	Q1 to Q4
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

$chart4->setTopLeftPosition('B40');
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
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$R$38:$V$38', NULL, 5),	//	2010
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
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$R$38:$V38', NULL, 5),	//	Q1 to Q4
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

$chart5->setTopLeftPosition('S40');
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
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$54:$E54', NULL, 5),	//	2010
	
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
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$A$54:$E54', NULL, 5),	//	Q1 to Q4
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

$chart6->setTopLeftPosition('B56');
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
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$H$54:$L54', NULL, 5),	//	2010
	
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
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$H$54:$L54', NULL, 5),	//	Q1 to Q4
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

$chart7->setTopLeftPosition('I56');
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
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$O$54:$S54', NULL, 5),	//	2010	
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
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$O$54:$S54', NULL, 5),	//	Q1 to Q4
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

$chart8->setTopLeftPosition('P56');
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
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$V$54:$Z$54', NULL, 5),	//	2010
	
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
	new PHPExcel_Chart_DataSeriesValues('String', 'Du_Pont_Analysis!$V$54:$Z54', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues9 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Du_Pont_Analysis!$V$55:$Z55', NULL, 5));

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

$chart9->setTopLeftPosition('W56');
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
$objPHPExcel->getActiveSheet()->getDefaultColumnDimension()->setWidth(5);
$objPHPExcel->getActiveSheet()->setShowGridlines(false);




$query2 = "select finyear, totalrevenue_sales, costrevenue,
grossprofit, operatingexp, operatinginc, 
incomebeforetax, netprofit
from income_statement where company='".$companyname."'";
$result = mysqli_query($conn,$query2);

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


while ($info = mysqli_fetch_array($result,MYSQLI_ASSOC))
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

$resultbalance = mysqli_query($conn,$query3);

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

while ($info1 = mysqli_fetch_array($resultbalance,MYSQLI_ASSOC))
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

$resultcash = mysqli_query($conn,$query4);

$results23 = array();
$results24 = array();
$results25 = array();

while ($info4 = mysqli_fetch_array($resultcash,MYSQLI_ASSOC))
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
$sheet->getStyle('A1:AD1')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->getStyle('A1:AD1')->getFill()->getStartColor()->setRGB('008000');

$objPHPExcel->getActiveSheet()->mergeCells('A1:AD1');
$objPHPExcel->getActiveSheet()->setCellValue('A1', 'Financial DashBoard');
$objPHPExcel->getActiveSheet()->getStyle("A1")->getFont()->setBold(true)
                                ->setName('Verdana')
                                ->setSize(20)
								->getColor()->setRGB('FFFFFF');

$objPHPExcel->getActiveSheet()->getStyle('A1:AD1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);


$objWorksheet->fromArray(
$results1,NULL,'A5'
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
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$A$5:$E$5', NULL, 5),
	
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker

$xAxisTickValues1 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$5:$E$5', NULL, 5),
	);	//	Q1 to Q4);
//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues1 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$A$6:E$6', NULL, 5),
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
$title1 = new PHPExcel_Chart_Title('');
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
$chart1->setTopLeftPosition('A5');
$chart1->setBottomRightPosition('F16');

$sheet->getStyle('A4:F4')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->getStyle('A4:F4')->getFill()->getStartColor()->setRGB('008000');

$objPHPExcel->getActiveSheet()->mergeCells('A4:F4');
$objPHPExcel->getActiveSheet()->setCellValue('A4', 'Total Revenue Sales');
$objPHPExcel->getActiveSheet()->getStyle("A4")->getFont()->setBold(true)
                                ->setName('Verdana')
                                ->setSize(8)
								->getColor()->setRGB('FFFFFF');

$objPHPExcel->getActiveSheet()->getStyle('A4:F4')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);


//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart1);

 //////////////////
 
//////////////////
 
//////////////////
 
//////////////////
 
//////////////////
 

$objPHPExcel->getActiveSheet()->fromArray($results2, NULL, 'G5');

$dataseriesLabels2 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$G$5', NULL, 5),	//	2010
	
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues2 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$G$5:$K$5', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues2 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$G$6:K$6', NULL, 5));

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
$title2 = new PHPExcel_Chart_Title('');
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
$chart2->setTopLeftPosition('G5');
$chart2->setBottomRightPosition('L16');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart2);

$sheet->getStyle('G4:L4')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->getStyle('G4:L4')->getFill()->getStartColor()->setRGB('008000');

$objPHPExcel->getActiveSheet()->mergeCells('G4:L4');
$objPHPExcel->getActiveSheet()->setCellValue('G4', 'Cost of Revenue/COGS');
$objPHPExcel->getActiveSheet()->getStyle("G4")->getFont()->setBold(true)
                                ->setName('Verdana')
                                ->setSize(8)
								->getColor()->setRGB('FFFFFF');

$objPHPExcel->getActiveSheet()->getStyle('G4:L4')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
$objPHPExcel->getActiveSheet()->fromArray($results3, NULL, 'M5');

$dataseriesLabels3 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$M$5:$Q5', NULL, 1),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues3 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$M$5:$Q$5', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues3 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$M$6:Q$6', NULL, 5));

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
$title3 = new PHPExcel_Chart_Title('');
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
$chart3->setTopLeftPosition('M5');
$chart3->setBottomRightPosition('R16');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart3);

$sheet->getStyle('M4:R4')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->getStyle('M4:R4')->getFill()->getStartColor()->setRGB('008000');

$objPHPExcel->getActiveSheet()->mergeCells('M4:R4');
$objPHPExcel->getActiveSheet()->setCellValue('M4', 'Cost of Revenue/COGS');
$objPHPExcel->getActiveSheet()->getStyle("M4")->getFont()->setBold(true)
                                ->setName('Verdana')
                                ->setSize(8)
								->getColor()->setRGB('FFFFFF');

$objPHPExcel->getActiveSheet()->getStyle('M4:R4')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
//////////////////////////////////////////////////
////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
$objPHPExcel->getActiveSheet()->fromArray($results4, NULL, 'S5');


$dataseriesLabels4 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$S$5:$W5', NULL, 5),	//	2010

);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues4 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$S$5:$W$5', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues4 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$S$6:W$6', NULL, 5),);

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
$title4 = new PHPExcel_Chart_Title('');
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
$chart4->setTopLeftPosition('S5');
$chart4->setBottomRightPosition('X16');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart4);

$sheet->getStyle('S4:X4')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->getStyle('S4:X4')->getFill()->getStartColor()->setRGB('008000');

$objPHPExcel->getActiveSheet()->mergeCells('S4:X4');
$objPHPExcel->getActiveSheet()->setCellValue('S4','Operating Expenses');
$objPHPExcel->getActiveSheet()->getStyle("S4")->getFont()->setBold(true)
                                ->setName('Verdana')
                                ->setSize(8)
								->getColor()->setRGB('FFFFFF');

$objPHPExcel->getActiveSheet()->getStyle('S4:X4')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
$objPHPExcel->getActiveSheet()->fromArray($results5, NULL, 'Y5');

$dataseriesLabels5 = array(
		new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$Y5:$AC5', NULL, 5),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues5 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$Y$5:$AC$5', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues5 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$Y$6:AC$6', NULL, 5),);

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
$title5 = new PHPExcel_Chart_Title('');
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
$chart5->setTopLeftPosition('Y5');
$chart5->setBottomRightPosition('AD16');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart5);

$sheet->getStyle('Y4:AD4')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->getStyle('Y4:AD4')->getFill()->getStartColor()->setRGB('008000');

$objPHPExcel->getActiveSheet()->mergeCells('Y4:AD4');
$objPHPExcel->getActiveSheet()->setCellValue('Y4','Operating Income');
$objPHPExcel->getActiveSheet()->getStyle("Y4")->getFont()->setBold(true)
                                ->setName('Verdana')
                                ->setSize(8)
								->getColor()->setRGB('FFFFFF');

$objPHPExcel->getActiveSheet()->getStyle('Y4:AD4')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////

$objPHPExcel->getActiveSheet()->fromArray($results6, NULL, 'A19');

$dataseriesLabels6 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$19:$E19', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues6 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$19:$E$19', NULL,5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues6 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$A$20:E$20', NULL, 5),);

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
$title6 = new PHPExcel_Chart_Title('');
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
$chart6->setTopLeftPosition('A19');
$chart6->setBottomRightPosition('F29');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart6);

$sheet->getStyle('A18:F18')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->getStyle('A18:F18')->getFill()->getStartColor()->setRGB('008000');

$objPHPExcel->getActiveSheet()->mergeCells('A18:F18');
$objPHPExcel->getActiveSheet()->setCellValue('A18','Income Before Tax');
$objPHPExcel->getActiveSheet()->getStyle("A18")->getFont()->setBold(true)
                                ->setName('Verdana')
                                ->setSize(8)
								->getColor()->setRGB('FFFFFF');

$objPHPExcel->getActiveSheet()->getStyle('A18:F18')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);



//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
$objPHPExcel->getActiveSheet()->fromArray($results7, NULL, 'G19');


$dataseriesLabels7 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$G$19:$K19', NULL, 5),	//	2010
	
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues7 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$G$19:$K$19', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues7 = array(
	 new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$G$20:K$20', NULL, 5),);

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
$title7 = new PHPExcel_Chart_Title('');
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
$chart7->setTopLeftPosition('G19');
$chart7->setBottomRightPosition('L29');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart7);

$sheet->getStyle('G18:L18')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->getStyle('G18:L18')->getFill()->getStartColor()->setRGB('008000');

$objPHPExcel->getActiveSheet()->mergeCells('G18:L18');
$objPHPExcel->getActiveSheet()->setCellValue('G18','Net Profit');
$objPHPExcel->getActiveSheet()->getStyle("G18")->getFont()->setBold(true)
                                ->setName('Verdana')
                                ->setSize(8)
								->getColor()->setRGB('FFFFFF');

$objPHPExcel->getActiveSheet()->getStyle('G18:L18')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
$objPHPExcel->getActiveSheet()->fromArray($results8, NULL, 'M19');

$dataseriesLabels8 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$M$19:$Q19', NULL, 5),	//	2010
	
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues8 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$M$19:$Q$19', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues8 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$M$20:Q$20', NULL, 5));

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
$title8 = new PHPExcel_Chart_Title('');
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
$chart8->setTopLeftPosition('M19');
$chart8->setBottomRightPosition('R29');
$objWorksheet->addChart($chart8);


$sheet->getStyle('M18:R18')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->getStyle('M18:R18')->getFill()->getStartColor()->setRGB('008000');

$objPHPExcel->getActiveSheet()->mergeCells('M18:R18');
$objPHPExcel->getActiveSheet()->setCellValue('M18','Cash and Equivalents');
$objPHPExcel->getActiveSheet()->getStyle("M18")->getFont()->setBold(true)
                                ->setName('Verdana')
                                ->setSize(8)
								->getColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('M18:R18')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
//////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////
$objPHPExcel->getActiveSheet()->fromArray($results9, NULL, 'S19');

$dataseriesLabels9 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$S$19:$W19', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues9 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$S$19:$W$19', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues9 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$S$20:W$20', NULL, 5),);

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
$title9 = new PHPExcel_Chart_Title('');
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
$chart9->setTopLeftPosition('S19');
$chart9->setBottomRightPosition('X29');
//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart9);

$sheet->getStyle('S18:X18')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->getStyle('S18:X18')->getFill()->getStartColor()->setRGB('008000');

$objPHPExcel->getActiveSheet()->mergeCells('S18:X18');
$objPHPExcel->getActiveSheet()->setCellValue('S18','Account Receivable');
$objPHPExcel->getActiveSheet()->getStyle("S18")->getFont()->setBold(true)
                                ->setName('Verdana')
                                ->setSize(8)
								->getColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('S18:X18')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

/////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////
$objPHPExcel->getActiveSheet()->fromArray($results10, NULL, 'Y19');

$dataseriesLabels10 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$Y$19:$AC19', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues10 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$Y$19:$AC$19', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues10 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$Y$20:AC$20', NULL, 5));

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
$title10 = new PHPExcel_Chart_Title('');
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
$chart10->setTopLeftPosition('Y19');
$chart10->setBottomRightPosition('AD29');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart10);

$sheet->getStyle('Y18:AD18')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->getStyle('Y18:AD18')->getFill()->getStartColor()->setRGB('008000');

$objPHPExcel->getActiveSheet()->mergeCells('Y18:AD18');
$objPHPExcel->getActiveSheet()->setCellValue('Y18','Inventory');
$objPHPExcel->getActiveSheet()->getStyle("Y18")->getFont()->setBold(true)
                                ->setName('Verdana')
                                ->setSize(8)
								->getColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('Y18:AD18')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);











//////////////////////////////
///////////////////////////////
//////////////////////////////
///////////////////////////////
//////////////////////////////
///////////////////////////////

$objPHPExcel->getActiveSheet()->fromArray($results11, NULL, 'A32');

$dataseriesLabels11 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$32:$E32', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues11 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$32:$E$32', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues11 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$A$33:$E$33', NULL, 5));

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
$title11 = new PHPExcel_Chart_Title('');
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
$chart11->setTopLeftPosition('A32');
$chart11->setBottomRightPosition('F42');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart11);
$sheet->getStyle('A31:F31')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->getStyle('A18:F31')->getFill()->getStartColor()->setRGB('008000');

$objPHPExcel->getActiveSheet()->mergeCells('A31:F31');
$objPHPExcel->getActiveSheet()->setCellValue('A31','Total Current Assets');
$objPHPExcel->getActiveSheet()->getStyle("A31")->getFont()->setBold(true)
                                ->setName('Verdana')
                                ->setSize(8)
								->getColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('A31:F31')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

//////////////////////////////
///////////////////////////////
//////////////////////////////
///////////////////////////////
//////////////////////////////
///////////////////////////////

$objPHPExcel->getActiveSheet()->fromArray($results12, NULL, 'G32');

$dataseriesLabels12 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$G$32:$K32', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues12 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$G$32:$K$32', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues12 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$G$33:$K33', NULL, 5));

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
$title12 = new PHPExcel_Chart_Title('');
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
$chart12->setTopLeftPosition('G32');
$chart12->setBottomRightPosition('L42');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart12);
$sheet->getStyle('G31:L31')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->getStyle('G31:L31')->getFill()->getStartColor()->setRGB('008000');

$objPHPExcel->getActiveSheet()->mergeCells('G31:L31');
$objPHPExcel->getActiveSheet()->setCellValue('G31','Property /Equipments');
$objPHPExcel->getActiveSheet()->getStyle("G31")->getFont()->setBold(true)
                                ->setName('Verdana')
                                ->setSize(8)
								->getColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('G31:L31')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
/////////////////////////////
///////////////////////////////
//////////////////////////////
///////////////////////////////
//////////////////////////////
///////////////////////////////


$objPHPExcel->getActiveSheet()->fromArray($results13, NULL, 'M32');

$dataseriesLabels13 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$M$32:$Q32', NULL, 5),	//	2010
	
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues13 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$M$32:$Q$32', NULL, 5),	//	Q1 to Q4);
);
//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues13 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$M$33:$Q33', NULL, 5),
);

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
$title13 = new PHPExcel_Chart_Title('');
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
$chart13->setTopLeftPosition('M32');
$chart13->setBottomRightPosition('R42');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart13);

$sheet->getStyle('M31:R31')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->getStyle('M31:R31')->getFill()->getStartColor()->setRGB('008000');

$objPHPExcel->getActiveSheet()->mergeCells('M31:R31');
$objPHPExcel->getActiveSheet()->setCellValue('M31','Long Term Investments');
$objPHPExcel->getActiveSheet()->getStyle("M31")->getFont()->setBold(true)
                                ->setName('Verdana')
                                ->setSize(8)
								->getColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('M31:R31')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

////////////////////////////////////////////////////////
///////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////

$objPHPExcel->getActiveSheet()->fromArray($results14, NULL, 'S32');

$dataseriesLabels14 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$S$32:$W32', NULL,5),	//	2010
	
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues14 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$S$32:$W$32', NULL,5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues14 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$S$33:$W33', NULL, 5),);

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
$title14 = new PHPExcel_Chart_Title('');
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
$chart14->setTopLeftPosition('S32');
$chart14->setBottomRightPosition('X42');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart14);
$sheet->getStyle('S31:X31')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->getStyle('S31:X31')->getFill()->getStartColor()->setRGB('008000');

$objPHPExcel->getActiveSheet()->mergeCells('S31:X31');
$objPHPExcel->getActiveSheet()->setCellValue('S31','Long Term Liabilities');
$objPHPExcel->getActiveSheet()->getStyle("S31")->getFont()->setBold(true)
                                ->setName('Verdana')
                                ->setSize(8)
								->getColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('S31:X31')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

/////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////


$objPHPExcel->getActiveSheet()->fromArray($results15, NULL, 'Y32');

$dataseriesLabels15 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$Y$32:$AC32', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues15 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$Y$32:$AC$32', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues15 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$Y$33:$AC33', NULL, 5),);

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
$title15 = new PHPExcel_Chart_Title('');
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
$chart15->setTopLeftPosition('Y32');
$chart15->setBottomRightPosition('AD42');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart15);

$sheet->getStyle('Y31:AD31')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->getStyle('Y31:AD31')->getFill()->getStartColor()->setRGB('008000');

$objPHPExcel->getActiveSheet()->mergeCells('Y31:AD31');
$objPHPExcel->getActiveSheet()->setCellValue('Y31','Total Assets');
$objPHPExcel->getActiveSheet()->getStyle("Y31")->getFont()->setBold(true)
                                ->setName('Verdana')
                                ->setSize(8)
								->getColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('Y31:AD31')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);


/////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////


$objPHPExcel->getActiveSheet()->fromArray($results16, NULL, 'A45');

$dataseriesLabels16 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$45:$E45', NULL, 5),	//	2010

);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues16 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$45:$E$45', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues16 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$A$46:E$46', NULL, 5),
	);

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
$title16 = new PHPExcel_Chart_Title('');
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

$objWorksheet->addChart($chart16);

//	Set the position where the chart should appear in the Financial_Dashboard
$chart16->setTopLeftPosition('A45');
$chart16->setBottomRightPosition('F55');
//	Add the chart to the Financial_Dashboard
$sheet->getStyle('A44:F44')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->getStyle('A44:F44')->getFill()->getStartColor()->setRGB('008000');

$objPHPExcel->getActiveSheet()->mergeCells('A44:F44');
$objPHPExcel->getActiveSheet()->setCellValue('A44','Accounts Payable');
$objPHPExcel->getActiveSheet()->getStyle("A44")->getFont()->setBold(true)
                                ->setName('Verdana')
                                ->setSize(8)
								->getColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('A44:F44')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

//////////////////////////////
///////////////////////////////
//////////////////////////////
///////////////////////////////
//////////////////////////////
///////////////////////////////

$objPHPExcel->getActiveSheet()->fromArray($results17, NULL, 'G45');

$dataseriesLabels17 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$G$45:$K45', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues17 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$G$45:$K$45', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues17 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$G$45:$K45', NULL, 5));

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
$title17 = new PHPExcel_Chart_Title('');
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
$chart17->setTopLeftPosition('G45');
$chart17->setBottomRightPosition('L55');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart17);
$sheet->getStyle('G44:L44')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->getStyle('G44:L44')->getFill()->getStartColor()->setRGB('008000');

$objPHPExcel->getActiveSheet()->mergeCells('G44:L44');
$objPHPExcel->getActiveSheet()->setCellValue('G44','Total Current Liabiliteis');
$objPHPExcel->getActiveSheet()->getStyle("G44")->getFont()->setBold(true)
                                ->setName('Verdana')
                                ->setSize(8)
								->getColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('G44:L44')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

//////////////////////////////
///////////////////////////////
//////////////////////////////
///////////////////////////////
//////////////////////////////
///////////////////////////////


$objPHPExcel->getActiveSheet()->fromArray($results18, NULL, 'M45');

$dataseriesLabels18 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$M$45:$Q45', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues18 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$M$45:$Q$45', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues18 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$M$46:$Q46', NULL, 5),);

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
$title18 = new PHPExcel_Chart_Title('');
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
$chart18->setTopLeftPosition('M45');
$chart18->setBottomRightPosition('R55');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart18);

$sheet->getStyle('M44:R44')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->getStyle('M44:R44')->getFill()->getStartColor()->setRGB('008000');

$objPHPExcel->getActiveSheet()->mergeCells('M44:R44');
$objPHPExcel->getActiveSheet()->setCellValue('M44','Long Term Debt');
$objPHPExcel->getActiveSheet()->getStyle("M44")->getFont()->setBold(true)
                                ->setName('Verdana')
                                ->setSize(8)
								->getColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('M44:R44')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);


////////////////////////////////////////////////////////
///////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////



$objPHPExcel->getActiveSheet()->fromArray($results19, NULL, 'S45');

$dataseriesLabels19 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$S$45', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues19 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$S$45:$W$45', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues19 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$S$45:$W45', NULL, 5));

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
$title19 = new PHPExcel_Chart_Title('');
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
$chart19->setTopLeftPosition('S45');
$chart19->setBottomRightPosition('X55');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart19);
$sheet->getStyle('S44:X44')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->getStyle('S44:X44')->getFill()->getStartColor()->setRGB('008000');

$objPHPExcel->getActiveSheet()->mergeCells('S44:X44');
$objPHPExcel->getActiveSheet()->setCellValue('S44','Long Term Liabiliteis');
$objPHPExcel->getActiveSheet()->getStyle("S44")->getFont()->setBold(true)
                                ->setName('Verdana')
                                ->setSize(8)
								->getColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('S44:X44')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

////////////////////////////////////////////////////////
///////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////
$objPHPExcel->getActiveSheet()->fromArray($results20, NULL, 'Y45');

$dataseriesLabels20 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$Y$45:$AC$45', NULL, 5),	//	2010
	
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues20 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$Y$45:$AC45', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues20 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$Y$46:$AC46', NULL, 5));

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
$title20 = new PHPExcel_Chart_Title('');
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
$chart20->setTopLeftPosition('Y45');
$chart20->setBottomRightPosition('AD55');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart20);

$sheet->getStyle('Y44:AD44')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->getStyle('Y44:AD44')->getFill()->getStartColor()->setRGB('008000');

$objPHPExcel->getActiveSheet()->mergeCells('Y44:AD44');
$objPHPExcel->getActiveSheet()->setCellValue('Y44','Total Liabiliteis');
$objPHPExcel->getActiveSheet()->getStyle("Y44")->getFont()->setBold(true)
                                ->setName('Verdana')
                                ->setSize(8)
								->getColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('Y44:AD44')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);


////////////////////////////////////////////////////////
///////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////

$objPHPExcel->getActiveSheet()->fromArray($results21, NULL, 'A58');

$dataseriesLabels21 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$58:$E58', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues21 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$A$58:$E58', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues21 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$A$59:$E59', NULL, 5));

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
$title21 = new PHPExcel_Chart_Title('');
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
$chart21->setTopLeftPosition('A58');
$chart21->setBottomRightPosition('F68');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart21);
$sheet->getStyle('A57:F57')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->getStyle('A57:F57')->getFill()->getStartColor()->setRGB('008000');

$objPHPExcel->getActiveSheet()->mergeCells('A57:F57');
$objPHPExcel->getActiveSheet()->setCellValue('A57','Total Equity');
$objPHPExcel->getActiveSheet()->getStyle("A57")->getFont()->setBold(true)
                                ->setName('Verdana')
                                ->setSize(8)
								->getColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('A57:F57')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);


////////////////////////////////////////////////////////
///////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////
$objPHPExcel->getActiveSheet()->fromArray($results22, NULL, 'G58');

$dataseriesLabels22 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$G$58:$K58', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues22 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$G$58:$K58', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues22 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$G$59:$K59', NULL, 5),);

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
$title22 = new PHPExcel_Chart_Title('');
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
$chart22->setTopLeftPosition('G58');
$chart22->setBottomRightPosition('L68');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart22);

$sheet->getStyle('G57:L57')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->getStyle('G57:L57')->getFill()->getStartColor()->setRGB('008000');

$objPHPExcel->getActiveSheet()->mergeCells('G57:L57');
$objPHPExcel->getActiveSheet()->setCellValue('G57','Liabilities & Equity');
$objPHPExcel->getActiveSheet()->getStyle("G57")->getFont()->setBold(true)
                                ->setName('Verdana')
                                ->setSize(8)
								->getColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('G57:L57')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);



////////////////////////////////////////////////////////
///////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////

$objPHPExcel->getActiveSheet()->fromArray($results23, NULL, 'M58');

$dataseriesLabels23 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$M$58:$Q58', NULL, 5),	//	2010
	
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues23 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$M$58:$Q58', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues23 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$M$59:$Q59', NULL, 5),);

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
$title23 = new PHPExcel_Chart_Title('');
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
$chart23->setTopLeftPosition('M58');
$chart23->setBottomRightPosition('R68');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart23);

$sheet->getStyle('M57:R57')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->getStyle('M57:R57')->getFill()->getStartColor()->setRGB('008000');

$objPHPExcel->getActiveSheet()->mergeCells('M57:R57');
$objPHPExcel->getActiveSheet()->setCellValue('M57','Cash from Operations');
$objPHPExcel->getActiveSheet()->getStyle("M57")->getFont()->setBold(true)
                                ->setName('Verdana')
                                ->setSize(8)
								->getColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('M57:R57')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
////////////////////////////////////////////////////////
///////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////

$objPHPExcel->getActiveSheet()->fromArray($results24, NULL, 'S58');

$dataseriesLabels24 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$S$58:$W58', NULL, 5),	//	2010
	
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues24 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$S$58:$W58', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues24 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$S$59:$W59', NULL, 5),);

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
$title24 = new PHPExcel_Chart_Title('');
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
$chart24->setTopLeftPosition('S58');
$chart24->setBottomRightPosition('X68');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart24);

$sheet->getStyle('S57:X57')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->getStyle('S57:X57')->getFill()->getStartColor()->setRGB('008000');

$objPHPExcel->getActiveSheet()->mergeCells('S57:X57');
$objPHPExcel->getActiveSheet()->setCellValue('S57','Cash From Investing');
$objPHPExcel->getActiveSheet()->getStyle("S57")->getFont()->setBold(true)
                                ->setName('Verdana')
                                ->setSize(8)
								->getColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('S57:X57')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);


////////////////////////////////////////////////////////
///////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////

$objPHPExcel->getActiveSheet()->fromArray($results25, NULL, 'Y58');

$dataseriesLabels25 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$Y$58:$AC58', NULL, 5),	//	2010
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues25 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Financial_Dashboard!$Y$58:$AC58', NULL, 5),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues25 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Financial_Dashboard!$Y$59:$AC59', NULL, 5));

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
$title25 = new PHPExcel_Chart_Title('');
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
$chart25->setTopLeftPosition('Y58');
$chart25->setBottomRightPosition('AD68');

//	Add the chart to the Financial_Dashboard
$objWorksheet->addChart($chart25);


$sheet->getStyle('Y57:AD57')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$sheet->getStyle('Y57:AD57')->getFill()->getStartColor()->setRGB('008000');

$objPHPExcel->getActiveSheet()->mergeCells('Y57:AD57');
$objPHPExcel->getActiveSheet()->setCellValue('Y57','Cash from Financing');
$objPHPExcel->getActiveSheet()->getStyle("Y57")->getFont()->setBold(true)
                                ->setName('Verdana')
                                ->setSize(8)
								->getColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('Y57:AD57')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
////////////////////////////////////////////////////////
///////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////

// Save Excel 2007 file
echo date('H:i:s') , " Write to Excel2007 format" , EOL;
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->setIncludeCharts(TRUE);
$file='outputfilefolder/output'.$companyname;
$file=$file.".xlsx";
$objWriter->save($file);
echo date('H:i:s') , " File written to " , str_replace('.php', '.xlsx', pathinfo($file, PATHINFO_BASENAME)) , EOL;
$queryuploadfile="update uploadfiledata set outputfilename='output".$companyname.".xlsx' where companyname='".$companyname."'"; 
mysqli_query($conn,$queryuploadfile);
// Echo memory peak usage
echo date('H:i:s') , " Peak memory usage: " , (memory_get_peak_usage(true) / 1024 / 1024) , " MB" , EOL;
// Echo done
echo date('H:i:s') , " Done writing file" , EOL;
echo 'File has been created in ' , getcwd() , EOL;
header('location:convertfile.php');

ob_flush();

?>