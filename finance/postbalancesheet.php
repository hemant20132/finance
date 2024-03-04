<?php
	error_reporting(E_ALL);

	$conn=mysqli_connect('fdb3.awardspace.net','2051670_db1','mastermagic2015','2051670_db1','3306','') or  die ("Error in query:".mysqli_error());
	
	$finyear=intval($_POST['finyear']);
	$company=$_POST['company'];
	$cashandequ=intval($_POST['cashandequivalents']);
	$accountrec=intval($_POST['accountreceivables']);
	$inventory=intval($_POST['inventory']);
	$prepaidexp=intval($_POST['prepaidexpenses']);
	$totcurassets=intval($_POST['totalcurrentassets']);
	$propertyequip=intval($_POST['propertyorequipment']);
	$goodwill=intval($_POST['goodwill']);
	$intangibles=intval($_POST['intangibles']);
	$longterminv=intval($_POST['longterminvestments']);
	$noterecelt=intval($_POST['notereceivablelongterm']);
	$othltassets=intval($_POST['otherlongtermassets']);
	$totltassets=intval($_POST['totallongtermassets']);
	$totalassets=intval($_POST['totalassets']);
		
	$acpayable=intval($_POST['accountpayable']);
	$payaccrue=intval($_POST['payableoraccrued']);
	$accruexp=intval($_POST['accruedexpenses']);
	$notpayableshort=intval($_POST['notespayableshort1']);
	$curportofltdebt=intval($_POST['currentportofltdebt']);
	$othercurliab=intval($_POST['othercurrentliabilities']);
	$totcurliab=intval($_POST['totalcurrentliabilities']);
	$ltdebpt=intval($_POST['longtermdebt']);
	$defitax=intval($_POST['deferredincometax']);
	$mininterest=intval($_POST['minorityinterest']);
	$othliab=intval($_POST['otherliabilities']);
	$ltliab=intval($_POST['longtermliabilities']);
	$totalliabilities=intval($_POST['totalliabilities']);
	
	$redemprefstock=intval($_POST['redemprefstock']);
	$prefstknonredem=intval($_POST['prefstknonredem']);
	$cmnstock=intval($_POST['cmnstock']);
	$retearn=intval($_POST['retearn']);
	$treasstk=intval($_POST['treasstk']);
	$unrealgain=intval($_POST['unrealgain']);
	$othequity=intval($_POST['othequity']);
	$totequity=intval($_POST['totequity']);
	$liabequity=intval($_POST['liabequity']);
	$totcomsharesout=intval($_POST['totcomsharesout']);
	$totprefsharesout=intval($_POST['totprefsharesout']);
	

	$query1 = "INSERT INTO balance_sheet(finyear, company, cashandequivalents, accountrecive,
	inventory, prepaidexp, totalcurrenassets, propertyequip, goodwill, intangibles, 
	longterminvst, notrecelongterm, otherlongterm, totallongterm, totalassets, accountpayable, 
	payable_accrued, accruedexp, notespayableshorttermdebt, curportltdeptcapleases, othcurrliab, 
	totalcurliab, longtermdebt, deferreditax, minorityint, otherliab, longtermliab,
	totalliabilities, redeemableprefstock, prefstocknonredeemable, commanstock, 
	retainearning, treasurystockcommon, unutilizedgain_loss, other_equity, totalequity, 
	liabilitiesandequity, totcomshareoutstand, totalprefshareoutstand)
	VALUES ("
	. $finyear . ",'" .$company. "'," . $cashandequ . "," . $accountrec . "," . $inventory .
	"," . $prepaidexp . "," . $totcurassets ."," . $propertyequip . "," . $goodwill .
	"," . $intangibles . "," . $longterminv . "," . $noterecelt . "," . $othltassets . "," . $totltassets  .
	"," . $totalassets  .  "," . $acpayable . "," . $payaccrue . "," . $accruexp . "," . $notpayableshort .
	"," . $curportofltdebt . "," . $othercurliab . "," . $totcurliab . "," . $ltdebpt . 
	"," . $defitax . "," . $mininterest . "," . $othliab . "," . $ltliab . "," . $totalliabilities .
	"," . $redemprefstock ."," . $prefstknonredem ."," . $cmnstock ."," . $retearn.
	"," .$treasstk . "," . $unrealgain .",". $othequity . "," . $totequity ."," . $liabequity.
	"," .$totcomsharesout ."," . $totprefsharesout . ")";

//mysqli_query($conn,$query1);
//mysqli_close($conn);
	mysqli_query($conn,$query1); 		
	mysqli_close($conn);

header("location:balancesheet.php?poststatus=record posted successfully");
	

	
?>