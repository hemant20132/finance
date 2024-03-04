<?php
set_include_path(get_include_path() . PATH_SEPARATOR . '../Classes/');
include 'PHPExcel.php';
include 'PHPExcel/DataSeriesValues.php';
include 'PHPExcel/DataSeries.php';
include 'PHPExcel/PlotArea.php';
include 'PHPExcel/Legend.php';
include 'PHPExcel/Title.php';
include 'PHPExcel/Chart.php';

	 
	$objPHPExcel = new PHPExcel();
	$objWorksheet = $objPHPExcel->getActiveSheet();
	//data for the worksheet graph
	$objWorksheet->fromArray(
	    array(
	        array('',   2010,   2011,   2012),
	        array('Q1',   12,   15,     21),
	        array('Q2',   56,   73,     86),
	        array('Q3',   52,   61,     69),
	        array('Q4',   30,   32,     0),
    ));
	
	
	$dataseriesLabels = array(
	    new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$B$1', null, 1),   //  2010
	    new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$C$1', null, 1),   //  2011
	    new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$D$1', null, 1),   //  2012
	);
	
	$xAxisTickValues = array(
    new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$2:$A$5', null, 4),  //  Q1 to Q4
	);

	$dataSeriesValues = array(
	    new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$B$2:$B$5', null, 4),
	    new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$C$2:$C$5', null, 4),
	    new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$D$2:$D$5', null, 4),
	);
	 
	$series = new PHPExcel_Chart_DataSeries(
	    PHPExcel_Chart_DataSeries::TYPE_BARCHART,      
	    PHPExcel_Chart_DataSeries::GROUPING_STANDARD,  
	    range(0, count($dataSeriesValues)-1),          
	    $dataseriesLabels,                              
	    $xAxisTickValues,                               
	    $dataSeriesValues                               
	);

	$series->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

	$plotarea = new PHPExcel_Chart_PlotArea(null, array($series));
	//  Set the chart legend
	$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, null, false);
	 
	$title = new PHPExcel_Chart_Title('Test Column Chart');
	$yAxisLabel = new PHPExcel_Chart_Title('Value ($k)');
	 
	//  Create the chart
	$chart = new PHPExcel_Chart(
	    'chart1',       // name
	    $title,         // title
	    $legend,        // legend
	    $plotarea,      // plotArea
		true,           // plotVisibleOnly
	    0,              // displayBlanksAs
	    null,           // xAxisLabel
	    $yAxisLabel     // yAxisLabel
	);
	 
	//  Set the position where the chart should appear in the worksheet
	$chart->setTopLeftPosition('A7');
	$chart->setBottomRightPosition('H20');
	 
	//  Add the chart to the worksheet
	$objWorksheet->addChart($chart);
	 
	 
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