<?php
session_start(); 
$companynm=$_SESSION['companyname'];
$inputfilename=$_SESSION['inputfilename'];

ob_start();

error_reporting(E_ALL);
ini_set("display_errors", 1);

$companynm = $_GET['companyname'];
$inputfilename = $_GET['inputfile'];

$inputfilename1="Excel/uploads/".$inputfilename;
$conn=mysqli_connect('sql211.epizy.com','epiz_33239466','Horizon2022','epiz_33239466_new','3306','') or  die ("Error in query:".mysqli_error());

echo $inputfilename1;

if ($companynm ="" or $inputfilename="")
{
$msg='no row selected';
header('location:convertfile.php?msg='.$msg);
}

require_once 'PHPExcel_1.8.0_doc/Classes/PHPExcel.php';
include ('PHPExcel_1.8.0_doc/Classes/PHPExcel/IOFactory.php');
$objReader = PHPExcel_IOFactory::createReader('Excel2007');
$objReader->setReadDataOnly(true);
$objPHPExcel = $objReader->load($inputfilename1);
$objReader->setLoadSheetsOnly( array("___Income_Statement___", "___Balance_Sheet___", "___Cash_Flow___") );
$objWorksheet=$objPHPExcel->setActiveSheetIndex(0);
$highestRow = $objWorksheet->getHighestRow(); // e.g. 10
$highestColumn = $objWorksheet->getHighestColumn(); // e.g 'F'
$highestColumnIndex = PHPExcel_Cell::columnIndexFromString($highestColumn); // e.g. 5

echo $highestColumnIndex;

$query1 = "update uploadfiledata set dataposted='YES' where companyname='". $companynm . "'";
mysqli_query($conn,$query1);

  $row = 1; 
  $companyname=array();
  for ($col=2;$col<=6;$col++) 
    	{
		$companyname[$col]=$objWorksheet->getCellByColumnAndRow(2,$row)->getCalculatedValue();
		}
		
  $row = 2; 
  $incfinyear=array();
  
  for ($col = 2; $col <= 6; $col++) 
       {
	$incfinyear[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
 $row = 6; 

$inctotrevsales=array();

    for ($col = 2; $col <= 6; $col++) 
  {
	$inctotrevsales[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
  }

 $row = 7; 
 
  
$costofrevenue=array();

    for ($col = 2; $col <= 6; $col++) 
  {
	$costofrevenue[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
  }
 
 $row = 8; 
   
$grossprofit=array();

    for ($col = 2; $col <= 6; $col++) 
  {
	$grossprofit[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
  }
 
  
$row = 12; 
  
$genadmexp=array();

    for ($col = 2; $col <= 6; $col++) 
  {
	$genadmexp[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
  }

$row = 13; 
  
$resev=array();

    for ($col = 2; $col <= 6; $col++) 
  {
	$resdev[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
  }

$row = 14; 
  
$salmarkexp=array();

    for ($col = 2; $col <= 6; $col++) 
  {
	$salmarkexp[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
  }

 $row = 15; 
  
$depamort=array();

    for ($col = 2; $col <= 6; $col++) 
  {
	$depamort[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
  }

 $row = 16; 
  
$intexp =array();

    for ($col = 2; $col <= 6; $col++) 
  {
	$intexp[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
  }

$row = 17; 
  
$specific1 =array();

    for ($col = 2; $col <= 6; $col++) 
  {
	$specific1[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
  }

$row = 18; 
  
$specific2 =array();

    for ($col = 2; $col <= 6; $col++) 
  {
	$specific2[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
  }

$row = 19; 
  
$specific3 =array();

    for ($col = 2; $col <= 6; $col++) 
  {
	$specific3[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
  }
  
$row = 20; 
  
$miscexp =array();

    for ($col = 2; $col <= 6; $col++) 
  {
	$miscexp[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
  }
  
$row=21;
$opexp =array();

    for ($col = 2; $col <= 6; $col++) 
  {
	$opexp[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
  }
  
  
$row=25;
$opinc =array();

    for ($col = 2; $col <= 6; $col++) 
  {
	$opinc[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
  }
  
  
$row=27;
$intexp2 =array();

    for ($col = 2; $col <= 6; $col++) 
  {
	$intexp2[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
  }
  
$row=28;
$incbetax =array();

    for ($col = 2; $col <= 6; $col++) 
  {
	$incbetax[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
  }

$row=30;
$tax =array();

    for ($col = 2; $col <= 6; $col++) 
  {
	$tax[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
  }
  
$row=31;
$netprofit =array();

    for ($col = 2; $col <= 6; $col++) 
  {
	$netprofit[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
  }
  


for ($col = 2; $col <= 6; $col++) 
  {
 $queryinsertinc="INSERT INTO income_statement (finyear, company, totalrevenue_sales, costrevenue, grossprofit,
 generaladmexpenses, researchdevelop, salesmarketingexp, depreciationamort, interestexp1, businespecific1, 
 businespecific2, businespecific3, miscexp, operatingexp, operatinginc, interestexp2,
 incomebeforetax, tax, netprofit) 
 VALUES (".$incfinyear[$col]. ",'" . $companyname[$col] . "'," .$inctotrevsales[$col] . "," . $costofrevenue[$col] . ",". $grossprofit[$col] 
 . "," . $genadmexp[$col] . "," . $resdev[$col] . "," .$salmarkexp[$col] . "," . $depamort[$col] . "," . $intexp[$col] . "," . $specific1[$col] 
  . ","  . $specific2[$col] . "," . $specific3[$col] . "," . $miscexp[$col] . "," . $opexp[$col] . "," . $opinc[$col] . "," . $intexp2[$col]
  ."," . $incbetax[$col] . "," . $tax[$col] . "," . $netprofit[$col] .")";
 echo $queryinsertinc;
 mysqli_query($conn,$queryinsertinc);
 }
	
	$objWorksheet=$objPHPExcel->setActiveSheetIndex(1);

  $row = 1; 
  $companyname1=array();
  for ($col=2;$col<=6;$col++) 
    	{
		$companyname1[$col]=$objWorksheet->getCellByColumnAndRow(2,$row)->getCalculatedValue();
		}
		
  $row = 2; 
  $balfinyear=array();
  
  for ($col = 2; $col <= 6; $col++) 
       {
	$balfinyear[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
 $row = 6; 
 $casheq=array();
  
  for ($col = 2; $col <= 6; $col++) 
       {
	$casheq[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
 
 $row = 7; 
 $accreceive =array();
  
  for ($col = 2; $col <= 6; $col++) 
       {
	$accreceive[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
 
 $row = 8; 
 $inventory =array();
  
  for ($col = 2; $col <= 6; $col++) 
       {
	$inventory[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
		
 $row = 9; 
 $prepaidexp =array();
  
  for ($col = 2; $col <= 6; $col++) 
       {
	$prepaidexp[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
 
 $row = 10; 
 $totcurassets =array();
  
  for ($col = 2; $col <= 6; $col++) 
       {
	$totcurassets[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
 
 
 $row = 12; 
 $propequip =array();
  
  for ($col = 2; $col <= 6; $col++) 
       {
	$propequip[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
 
 $row = 13; 
 $goodwill =array();
  
  for ($col = 2; $col <= 6; $col++) 
       {
	$goodwill[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
 
 $row = 14; 
 $intang =array();
  
  for ($col = 2; $col <= 6; $col++) 
       {
	$intang[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
 
 
 $row = 15; 
 $longterminvst =array();
  
  for ($col = 2; $col <= 6; $col++) 
       {
	$longterminvst[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
 
 $row = 16; 
 $notreceive =array();
  
  for ($col = 2; $col <= 6; $col++) 
       {
	$notreceive[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
 
 $row = 17; 
 $otherlongterm =array();
  
  for ($col = 2; $col <= 6; $col++) 
       {
	$otherlongterm[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
 
 $row = 18; 
 $totallongterm =array();
  
  for ($col = 2; $col <= 6; $col++) 
       {
	$totallongterm[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
 
 
 $row = 20; 
 $totalassets =array();
  
  for ($col = 2; $col <= 6; $col++) 
       {
	$totalassets[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
 
 $row = 24; 
 $accpay =array();
	
  for ($col = 2; $col <= 6; $col++) 
       {
	$accpay[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
 
 
 $row = 25; 
 $payacc =array();
	
  for ($col = 2; $col <= 6; $col++) 
       {
	$payacc[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
 
 
 
 $row = 26; 
 $accexp =array();
	
  for ($col = 2; $col <= 6; $col++) 
       {
	$accexp[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
 
$row = 27; 
 $notespay =array();
	
  for ($col = 2; $col <= 6; $col++) 
       {
	$notespay[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
 
$row = 28; 
 $caplease =array();
	
  for ($col = 2; $col <= 6; $col++) 
       {
	$caplease[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
 
$row = 29; 
 $otherliab =array();
	
  for ($col = 2; $col <= 6; $col++) 
       {
	$otherliab[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}

$row = 30; 
 $totalcurliab =array();
	
  for ($col = 2; $col <= 6; $col++) 
       {
	$totalcurliab[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
 
 
$row = 32; 
 $longtermdebt =array();
	
  for ($col = 2; $col <= 6; $col++) 
       {
	$longtermdebt[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
 
$row = 33; 
 $defitax =array();
	
  for ($col = 2; $col <= 6; $col++) 
       {
	$defitax[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
 
$row = 34; 
 $mininterest =array();
	
  for ($col = 2; $col <= 6; $col++) 
       {
	$mininterest[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
 
 
$row = 35; 
 $othliab =array();
		
  for ($col = 2; $col <= 6; $col++) 
       {
	$othliab[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
 
$row = 36; 
 $lngtliab  =array();
		
  for ($col = 2; $col <= 6; $col++) 
       {
	$lngtliab[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
 

$row = 38; 
 $totalliab  =array();
		
  for ($col = 2; $col <= 6; $col++) 
       {
	$totalliab[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}

$row = 42; 
 $reedmprfstk  =array();
		
  for ($col = 2; $col <= 6; $col++) 
       {
	$reedmprfstk[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
 
$row = 43; 
 $prfstk  =array();
		
  for ($col = 2; $col <= 6; $col++) 
       {
	$prfstk[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
 

$row = 44; 
 $cmnstk  =array();
		
  for ($col = 2; $col <= 6; $col++) 
       {
	$cmnstk[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
 
 
$row = 45; 
 $retearn  =array();
		
  for ($col = 2; $col <= 6; $col++) 
       {
	$retearn[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
 
 
$row = 46; 
 $tsrstock  =array();
		
  for ($col = 2; $col <= 6; $col++) 
       {
	$tsrstock[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
 

 $row = 47; 
 $unregain  =array();
		
  for ($col = 2; $col <= 6; $col++) 
       {
	$unregain[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}

 
 $row = 48; 
 $othequity  =array();
		
  for ($col = 2; $col <= 6; $col++) 
       {
	$othequity[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
		
$row = 49; 
 $totequity  =array();
		
  for ($col = 2; $col <= 6; $col++) 
       {
	$totequity[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
		
		
$row = 51; 
 $liabeqty =array();
		
  for ($col = 2; $col <= 6; $col++) 
       {
	$liabeqty[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
		
$row = 53; 
 $totcmnshrs =array();
		
  for ($col = 2; $col <= 6; $col++) 
       {
	$totcmnshrs[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}

$row = 54; 
 $totprfshrs =array();
		
  for ($col = 2; $col <= 6; $col++) 
       {
	$totprfshrs[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
		

for ($col = 2; $col <= 6; $col++) 
  {
 $queryinsertbal= "INSERT INTO balance_sheet (finyear, company, cashandequivalents, accountrecive, 
 inventory, prepaidexp, totalcurrenassets, propertyequip, goodwill, intangibles, 
 longterminvst, notrecelongterm, otherlongterm, totallongterm, totalassets, 
 accountpayable, payable_accrued, accruedexp, notespayableshorttermdebt, 
 curportltdeptcapleases, othcurrliab, totalcurliab, longtermdebt, deferreditax, 
 minorityint, otherliab, longtermliab, totalliabilities, redeemableprefstock, 
 prefstocknonredeemable, commanstock, retainearning, treasurystockcommon, unutilizedgain_loss, 
 other_equity, totalequity, liabilitiesandequity, totcomshareoutstand, totalprefshareoutstand) 
 VALUES (" .$balfinyear[$col] . ",'". $companyname1[$col] ."'," . $casheq[$col] . "," . $accreceive[$col] . ",".
$inventory[$col] .",". $prepaidexp[$col] .",". $totcurassets[$col] ."," . $propequip[$col] .",". $goodwill[$col] .",".$intang[$col] .",".
$longterminvst[$col] .",". $notreceive[$col] .",".$otherlongterm[$col] ."," . $totallongterm[$col] . ",". $totalassets[$col] .",".
$accpay[$col] .",". $payacc[$col] .",". $accexp[$col] .",". $notespay[$col] ."," . 
$caplease[$col] .",". $otherliab[$col] . "," . $totalcurliab[$col] . "," . $longtermdebt[$col] .",". $defitax[$col] . "," .
$mininterest[$col]  . ",". $othliab[$col] . "," .$lngtliab[$col] . "," . $totalliab[$col] . "," .$reedmprfstk[$col] . "," .
$prfstk[$col] ."," .$cmnstk[$col] .",". $retearn[$col] .",".  $tsrstock[$col] .",".  $unregain[$col] .",".
$othequity[$col] .",". $totequity[$col] . "," . $liabeqty[$col] . "," . $totcmnshrs[$col]. ",". $totprfshrs[$col] .")";


mysqli_query($conn,$queryinsertbal);
}



 $objWorksheet=$objPHPExcel->setActiveSheetIndex(2);

  $row = 1; 
  $companyname2 = array();
  for ($col=2;$col<=6;$col++) 
    	{
		$companyname2[$col]=$objWorksheet->getCellByColumnAndRow(2,$row)->getCalculatedValue();
		}
		
  $row = 2; 
  $cashfinyear=array();
  
  for ($col = 2; $col <= 6; $col++) 
       {
	$cashfinyear[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
 $row = 6; 

$netincome=array();

  for ($col = 2; $col <= 6; $col++) 
       {
	$netincome[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
 $row = 7; 
 
$depdepla=array();

  for ($col = 2; $col <= 6; $col++) 
       {
	$depdepla[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}

$row = 8; 
 
$amort=array();

  for ($col = 2; $col <= 6; $col++) 
       {
	$amort[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}

$row = 9; 
 
$deftax =array();

  for ($col = 2; $col <= 6; $col++) 
       {
	$deftax[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}

$row = 10; 
 
$ncitems =array();

  for ($col = 2; $col <= 6; $col++) 
       {
	$ncitems[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}

$row = 11; 
 
$cworkcap =array();

  for ($col = 2; $col <= 6; $col++) 
       {
	$cworkcap[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}

$row = 12; 
 
$accrec =array();

  for ($col = 2; $col <= 6; $col++) 
       {
	$accrec[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
		
$row = 13; 
 
$inventory	 =array();

  for ($col = 2; $col <= 6; $col++) 
       {
	$inventory[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
		
$row = 14; 
 
$accpay	 =array();

  for ($col = 2; $col <= 6; $col++) 
       {
	$accpay[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
		
$row = 15; 
 
$othliab	 =array();

  for ($col = 2; $col <= 6; $col++) 
       {
	$othliab[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
		
$row = 16; 
 
$othopcashflow	 =array();

  for ($col = 2; $col <= 6; $col++) 
       {
	$othopcashflow[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
		
$row = 17; 
 
$cashfromop	 =array();

  for ($col = 2; $col <= 6; $col++) 
       {
	$cashfromop[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}

$row = 21; 
 
$purfixassets	 =array();

  for ($col = 2; $col <= 6; $col++) 
       {
	$purfixassets[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
		
$row = 22; 
 
$acofbusi	 =array();

  for ($col = 2; $col <= 6; $col++) 
       {
	$acofbusi[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
		
$row = 23; 
 
$salofbusi	 =array();

  for ($col = 2; $col <= 6; $col++) 
       {
	$salofbusi[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
		
$row = 24; 
		
$saloffixass	 =array();

  for ($col = 2; $col <= 6; $col++) 
       {
	$saloffixass[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}

$row = 25; 
		
$othinvcashflow	 =array();

  for ($col = 2; $col <= 6; $col++) 
       {
	$othinvcashflow[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
		
$row = 26; 
		
$cashfrminvst	 =array();

  for ($col = 2; $col <= 6; $col++) 
       {
	$cashfrminvst[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
		
$row = 30; 

$fincashflow	 =array();

  for ($col = 2; $col <= 6; $col++) 
       {
	$fincashflow[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
		
$row = 31; 

$totcshdivpaid	 =array();

  for ($col = 2; $col <= 6; $col++) 
       {
	$totcshdivpaid[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
		
		
$row = 32; 
$isofretstk	 =array();


  for ($col = 2; $col <= 6; $col++) 
       {
	$isofretstk[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}


$row = 33; 


$issuofdebt =array();

  for ($col = 2; $col <= 6; $col++) 
       {
	$issuofdebt[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
		
	
$row = 34; 

$cashfromfin  =array();

  for ($col = 2; $col <= 6; $col++) 
       {
	$cashfromfin[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
	
		
$row = 36; 

$netchgcash  =array();

  for ($col = 2; $col <= 6; $col++) 
       {
	$netchgcash[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
	
$row = 38; 

$netcashbeg  =array();

  for ($col = 2; $col <= 6; $col++) 
       {
	$netcashbeg[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
		
$row = 39; 

$netcashend  =array();

  for ($col = 2; $col <= 6; $col++) 
       {
	$netcashend[$col]=$objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
		}
		

for ($col = 2; $col <= 6; $col++) 
  {
 $queryinsertcash= 
 "INSERT INTO cash_flow(company, finyear, netincomestarting, depletion, amortization,
 deferred_taxes, noncashitems, changeinworkcap, accountreceiv, inventories,
 accountpayable, otherliab, otheropcashflow, cashfromoper, purchasefixassets, 
 acqofbusiness, saleofbusiness, saleoffixedassets, otherinvestcashflow, cashfrominvesting, 
 fincashflow, totalcashdivpaid, issueofstock, inssueofdebt, cashfromfinance, 
 netchangeincash, netincashbeginbal, netincashendbal) VALUES
 ('" . $companyname2[$col] . "'," . $cashfinyear[$col] . ","  . $netincome[$col] . ",". $depdepla[$col] . ",". $amort[$col] . "," .
 $deftax[$col] . "," . $ncitems[$col] . "," .$cworkcap[$col] . "," .  $accrec[$col] . "," . $inventory[$col] . ",".
 $accpay[$col] ."," . $othliab[$col] .",". $othopcashflow[$col] .",". $cashfromop[$col] . "," . $purfixassets[$col]  .",".
 $acofbusi[$col] ."," . $salofbusi[$col] . "," . $saloffixass[$col] . "," . $othinvcashflow[$col] . "," . $cashfrminvst[$col] .",".
 $fincashflow[$col] . "," . $totcshdivpaid[$col] .",". $isofretstk[$col] .",". $issuofdebt[$col] .",".  $cashfromfin[$col] .",".
 $netchgcash[$col] .",". $netcashbeg[$col] ."," .$netcashend[$col] . ")";
mysqli_query($conn,$queryinsertcash);
}


header("location:output.php?companyname=". $companynm ."& inputfile=". $inputfilename );
ob_flush();

?>