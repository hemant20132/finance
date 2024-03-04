<?php
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  header('Content-Disposition: attachment;filename="graphdemo.xlsx"');
  header('Cache-Control: max-age=0');

  $workbook->setActiveSheetIndex(0);
  $sheet = $workbook->getActiveSheet();
  $sheet->getPageMargins()->setTop(0.6);
  $sheet->getPageMargins()->setBottom(0.6);
  $sheet->getPageMargins()->setHeader(0.4);
  $sheet->getPageMargins()->setFooter(0.4);
  $sheet->getPageMargins()->setLeft(0.4);
  $sheet->getPageMargins()->setRight(0.4);
  $workbook->getProperties()->setTitle("Demo");
  $workbook->getProperties()->setCreator("Demo");
  $workbook->getProperties()->setLastModifiedBy("Demo");
  $workbook->getProperties()->setCompany("Demo");

  $data = array('10%', '20%', '30%', '40%', '50%', '60%', '70%', '80%', '90%', '100%');
  $row = 1;
  foreach($data as $point) {
    $sheet->setCellValueByColumnAndRow(0, $row++, $point);
  }

  $data = array(12, 56, 89, 45, 42, 22, 15, 8, 2, 0);
  $row = 1;
  foreach($data as $point) {
    $sheet->setCellValueByColumnAndRow(1, $row++, $point);
  }

  $values = new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$B$1:$B$10');
  $categories = new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$1:$A$10');

  $series = new PHPExcel_Chart_DataSeries(
    PHPExcel_Chart_DataSeries::TYPE_BARCHART,       // plotType
    PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,  // plotGrouping
    array(0),                                       // plotOrder
    array(),                                        // plotLabel
    array($categories),                             // plotCategory
    array($values)                                  // plotValues
  );
  $series->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);

  $layout = new PHPExcel_Chart_Layout();
  $plotarea = new PHPExcel_Chart_PlotArea($layout, array($series));

  $chart = new PHPExcel_Chart('sample', null, null, $plotarea);

  $chart->setTopLeftPosition('C1');
  $chart->setBottomRightPosition('J15');

  $sheet->addChart($chart);

  $writer = PHPExcel_IOFactory::createWriter($workbook, 'Excel2007');
  $writer->setIncludeCharts(TRUE);
  $writer->save('php://output');
  ?>