$sheetId=1;
$objPHPExcel->createSheet(NULL, $sheetId);
$objWorksheet=$objPHPExcel->setActiveSheetIndex(1);  
//$objPHPExcel->getActiveSheet()->setTitle('Du-Pont-Analysis');



$query1 = "select income_statement.finyear, income_statement.totalrevenue_sales,income_statement.costrevenue,
income_statement.netprofit, 
balance_sheet.totalassets, balance_sheet.inventory, balance_sheet.accountrecive, balance_sheet.totalequity from income_statement, 
balance_sheet where income_statement.company='".$companyname."' and income_statement.finyear=balance_sheet.finyear ";

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

//	Add the chart to the 
$objWorksheet->addChart($chart9);

/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////
/////////////////////////////////////

$sheetId=2;
$objPHPExcel->createSheet(NULL, $sheetId);
$objWorksheet=$objPHPExcel->setActiveSheetIndex(2);  
$objPHPExcel->getActiveSheet()->setTitle('Financial DashBoard');


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
  
  
 

$objPHPExcel->getActiveSheet()->mergeCells('A1:T1');
$objPHPExcel->getActiveSheet()->setCellValue('A1', 'Financial DashBoard');
$objPHPExcel->getActiveSheet()->getStyle("A1")->getFont()->setBold(true)
                                ->setName('Verdana')
                                ->setSize(20)
								->getColor()->setRGB('008000');

$objPHPExcel->getActiveSheet()->getStyle('A1:T1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);


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

$xAxisTickValues1 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$A$4:$E$4', NULL, NULL));	//	Q1 to Q4);
//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues1 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$A$5:E$5', NULL, 5));

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

//	Set the position where the chart should appear in the worksheet
$chart1->setTopLeftPosition('A4');
$chart1->setBottomRightPosition('F21');

//	Add the chart to the worksheet
$objWorksheet->addChart($chart1);

/////////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////////////

$objPHPExcel->getActiveSheet()->fromArray($results2, NULL, 'G1');


$dataseriesLabels2 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$G$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$G$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$G$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$G$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$G$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues2 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$G$2:$K$2', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues2 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$G$2:K$2', NULL, 5));

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

//	Set the position where the chart should appear in the worksheet
$chart2->setTopLeftPosition('G1');
$chart2->setBottomRightPosition('L17');

//	Add the chart to the worksheet
$objWorksheet->addChart($chart2);
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
$objPHPExcel->getActiveSheet()->fromArray($results3, NULL, 'M1');

$dataseriesLabels3 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$M$1', NULL, 1),	//	2010
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$M$2', NULL, 1),	//	2011
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$M$3', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$M$4', NULL, 1),	//	2012
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$M$5', NULL, 1),	//	2012
);

//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues3 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$M$1:$Q$1', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues3 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$M$2:Q$2', NULL, 5));

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

//	Set the position where the chart should appear in the worksheet
$chart3->setTopLeftPosition('M1');
$chart3->setBottomRightPosition('R17');

//	Add the chart to the worksheet
$objWorksheet->addChart($chart3);
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
$objPHPExcel->getActiveSheet()->fromArray($results4, NULL, 'S1');


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
$xAxisTickValues4 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$S$1:$W$1', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues4 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$S$2:W$2', NULL, 5));

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

//	Set the position where the chart should appear in the worksheet
$chart4->setTopLeftPosition('S1');
$chart4->setBottomRightPosition('X17');

//	Add the chart to the worksheet
$objWorksheet->addChart($chart4);
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
$objPHPExcel->getActiveSheet()->fromArray($results5, NULL, 'Y1');


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
$xAxisTickValues5 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$Y$1:$AC$1', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues5 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$Y$2:AC$2', NULL, 5));

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

//	Set the position where the chart should appear in the worksheet
$chart5->setTopLeftPosition('Y1');
$chart5->setBottomRightPosition('AD17');

//	Add the chart to the worksheet
$objWorksheet->addChart($chart5);
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////

$objPHPExcel->getActiveSheet()->fromArray($results6, NULL, 'A20');


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
$xAxisTickValues6 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$20:$F$20', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues6 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$A$21:F$21', NULL, 5));

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

//	Set the position where the chart should appear in the worksheet
$chart6->setTopLeftPosition('A20');
$chart6->setBottomRightPosition('F37');

//	Add the chart to the worksheet
$objWorksheet->addChart($chart6);
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
$objPHPExcel->getActiveSheet()->fromArray($results7, NULL, 'G20');


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
$xAxisTickValues7 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$G$20:$L$20', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues7 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$G$21:L$21', NULL, 5));

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

//	Set the position where the chart should appear in the worksheet
$chart7->setTopLeftPosition('G20');
$chart7->setBottomRightPosition('L37');

//	Add the chart to the worksheet
$objWorksheet->addChart($chart7);
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
$objPHPExcel->getActiveSheet()->fromArray($results8, NULL, 'M20');

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
$xAxisTickValues8 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$M$20:$Q$20', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues8 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$M$21:Q$21', NULL, 5));

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

//	Set the position where the chart should appear in the worksheet
$chart8->setTopLeftPosition('M20');
$chart8->setBottomRightPosition('R37');

//	Add the chart to the worksheet
$objWorksheet->addChart($chart8);
//////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////
$objPHPExcel->getActiveSheet()->fromArray($results9, NULL, 'S20');

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
$xAxisTickValues9 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$S$20:$W$20', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues9 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$S$21:W$21', NULL, 5));

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

//	Set the position where the chart should appear in the worksheet
$chart9->setTopLeftPosition('S20');
$chart9->setBottomRightPosition('X37');

//	Add the chart to the worksheet
$objWorksheet->addChart($chart9);


/////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////
$objPHPExcel->getActiveSheet()->fromArray($results10, NULL, 'Y20');

$dataseriesLabels10 = array(
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
$xAxisTickValues10 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$Y$20:$AC$20', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues10 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$Y$21:AC$21', NULL, 5));

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

//	Set the position where the chart should appear in the worksheet
$chart10->setTopLeftPosition('Y20');
$chart10->setBottomRightPosition('AD37');

//	Add the chart to the worksheet
$objWorksheet->addChart($chart10);

//////////////////////////////
///////////////////////////////
//////////////////////////////
///////////////////////////////
//////////////////////////////
///////////////////////////////

$objPHPExcel->getActiveSheet()->fromArray($results11, NULL, 'A40');

$dataseriesLabels11 = array(
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
$xAxisTickValues11 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$40:$E$40', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues11 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$A$21:E$21', NULL, 5));

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

//	Set the position where the chart should appear in the worksheet
$chart11->setTopLeftPosition('A40');
$chart11->setBottomRightPosition('F57');

//	Add the chart to the worksheet
$objWorksheet->addChart($chart11);

//////////////////////////////
///////////////////////////////
//////////////////////////////
///////////////////////////////
//////////////////////////////
///////////////////////////////

$objPHPExcel->getActiveSheet()->fromArray($results12, NULL, 'G40');

$dataseriesLabels12 = array(
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
$xAxisTickValues12 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$G$20:$K$20', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues12 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$G$21:$K21', NULL, 5));

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

//	Set the position where the chart should appear in the worksheet
$chart12->setTopLeftPosition('G40');
$chart12->setBottomRightPosition('L57');

//	Add the chart to the worksheet
$objWorksheet->addChart($chart12);

//////////////////////////////
///////////////////////////////
//////////////////////////////
///////////////////////////////
//////////////////////////////
///////////////////////////////


$objPHPExcel->getActiveSheet()->fromArray($results13, NULL, 'M40');

$dataseriesLabels13 = array(
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
$xAxisTickValues13 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$M$20:$R$20', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues13 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$M$21:$R21', NULL, 5));

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

//	Set the position where the chart should appear in the worksheet
$chart13->setTopLeftPosition('M40');
$chart13->setBottomRightPosition('R57');

//	Add the chart to the worksheet
$objWorksheet->addChart($chart13);

////////////////////////////////////////////////////////
///////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////

$objPHPExcel->getActiveSheet()->fromArray($results14, NULL, 'S40');

$dataseriesLabels14 = array(
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
$xAxisTickValues14 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$S$20:$W$20', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues14 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$S$21:$W21', NULL, 5));

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

//	Set the position where the chart should appear in the worksheet
$chart14->setTopLeftPosition('S40');
$chart14->setBottomRightPosition('X57');

//	Add the chart to the worksheet
$objWorksheet->addChart($chart14);

/////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////


$objPHPExcel->getActiveSheet()->fromArray($results15, NULL, 'Y40');

$dataseriesLabels15 = array(
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
$xAxisTickValues15 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$Y$20:$AC$20', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues15 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$Y$21:$AC21', NULL, 5));

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

//	Set the position where the chart should appear in the worksheet
$chart15->setTopLeftPosition('Y40');
$chart15->setBottomRightPosition('AD57');

//	Add the chart to the worksheet
$objWorksheet->addChart($chart15);

/////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////


$objPHPExcel->getActiveSheet()->fromArray($results16, NULL, 'A60');

$dataseriesLabels16 = array(
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
$xAxisTickValues16 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$60:$E$60', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues16 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$A$61:E$61', NULL, 5));

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

//	Set the position where the chart should appear in the worksheet
$chart16->setTopLeftPosition('A60');
$chart16->setBottomRightPosition('F77');

//	Add the chart to the worksheet
$objWorksheet->addChart($chart16);

//////////////////////////////
///////////////////////////////
//////////////////////////////
///////////////////////////////
//////////////////////////////
///////////////////////////////

$objPHPExcel->getActiveSheet()->fromArray($results17, NULL, 'G60');

$dataseriesLabels17 = array(
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
$xAxisTickValues17 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$G$60:$K$60', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues17 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$G$61:$K61', NULL, 5));

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

//	Set the position where the chart should appear in the worksheet
$chart17->setTopLeftPosition('G60');
$chart17->setBottomRightPosition('L77');

//	Add the chart to the worksheet
$objWorksheet->addChart($chart17);

//////////////////////////////
///////////////////////////////
//////////////////////////////
///////////////////////////////
//////////////////////////////
///////////////////////////////


$objPHPExcel->getActiveSheet()->fromArray($results18, NULL, 'M60');

$dataseriesLabels18 = array(
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
$xAxisTickValues18 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$M$60:$R$60', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues18 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$M$61:$R61', NULL, 5));

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

//	Set the position where the chart should appear in the worksheet
$chart18->setTopLeftPosition('M60');
$chart18->setBottomRightPosition('R77');

//	Add the chart to the worksheet
$objWorksheet->addChart($chart18);

////////////////////////////////////////////////////////
///////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////



$objPHPExcel->getActiveSheet()->fromArray($results19, NULL, 'S60');

$dataseriesLabels19 = array(
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
$xAxisTickValues19 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$S$60:$X$60', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues19 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$S$61:$X61', NULL, 5));

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

//	Set the position where the chart should appear in the worksheet
$chart19->setTopLeftPosition('S60');
$chart19->setBottomRightPosition('X77');

//	Add the chart to the worksheet
$objWorksheet->addChart($chart19);

////////////////////////////////////////////////////////
///////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////
$objPHPExcel->getActiveSheet()->fromArray($results20, NULL, 'Y60');

$dataseriesLabels20 = array(
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
$xAxisTickValues20 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$Y$60:$AC60', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues20 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$Y$61:$AC61', NULL, 5));

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

//	Set the position where the chart should appear in the worksheet
$chart20->setTopLeftPosition('Y60');
$chart20->setBottomRightPosition('AD77');

//	Add the chart to the worksheet
$objWorksheet->addChart($chart20);

////////////////////////////////////////////////////////
///////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////

$objPHPExcel->getActiveSheet()->fromArray($results21, NULL, 'A80');

$dataseriesLabels21 = array(
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
$xAxisTickValues21 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$80:$F$80', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues21 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$A$81:$F$81', NULL, 5));

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

//	Set the position where the chart should appear in the worksheet
$chart21->setTopLeftPosition('A80');
$chart21->setBottomRightPosition('F97');

//	Add the chart to the worksheet
$objWorksheet->addChart($chart21);

////////////////////////////////////////////////////////
///////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////
$objPHPExcel->getActiveSheet()->fromArray($results22, NULL, 'G80');

$dataseriesLabels22 = array(
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
$xAxisTickValues22 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$G$80:$L80', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues22 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$G$81:$L81', NULL, 5));

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

//	Set the position where the chart should appear in the worksheet
$chart22->setTopLeftPosition('G80');
$chart22->setBottomRightPosition('L97');

//	Add the chart to the worksheet
$objWorksheet->addChart($chart22);

////////////////////////////////////////////////////////
///////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////

$objPHPExcel->getActiveSheet()->fromArray($results23, NULL, 'M80');

$dataseriesLabels23 = array(
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
$xAxisTickValues23 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$M$80:$R80', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues23 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$M$81:$R81', NULL, 5));

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

//	Set the position where the chart should appear in the worksheet
$chart23->setTopLeftPosition('M80');
$chart23->setBottomRightPosition('R97');

//	Add the chart to the worksheet
$objWorksheet->addChart($chart23);

////////////////////////////////////////////////////////
///////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////

$objPHPExcel->getActiveSheet()->fromArray($results24, NULL, 'S80');

$dataseriesLabels24 = array(
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
$xAxisTickValues24 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$S$80:$W80', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues24 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$S$81:$W81', NULL, 5));

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

//	Set the position where the chart should appear in the worksheet
$chart24->setTopLeftPosition('S80');
$chart24->setBottomRightPosition('X97');

//	Add the chart to the worksheet
$objWorksheet->addChart($chart24);

////////////////////////////////////////////////////////
///////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////

$objPHPExcel->getActiveSheet()->fromArray($results25, NULL, 'Y80');

$dataseriesLabels25 = array(
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
$xAxisTickValues25 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$Y$81:$AC80', NULL, NULL),	//	Q1 to Q4
);

//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in	 series
//		Data values
//		Data Marker
$dataSeriesValues25 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$Y$81:$AC81', NULL, 5));

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

//	Set the position where the chart should appear in the worksheet
$chart25->setTopLeftPosition('Y80');
$chart25->setBottomRightPosition('AD97');

//	Add the chart to the worksheet
$objWorksheet->addChart($chart25);

////////////////////////////////////////////////////////
///////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//////////////////////////////////////////////////////

	