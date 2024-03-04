<?php
	error_reporting(E_ALL);

	$conn= mysqli_connect("localhost","root","","finance") or die(mysql_error());

//	$conn = mysql_connect("localhost","nacheeso_hemant","7ByV+sA[CsL6") or die(mysql_error());
//	mysql_select_db("nacheeso_finance");
	
	$finyear=intval($_POST['finyear']);
	$company=$_POST['company'];
	$netincomestarting=$_POST['netincomestartingline'];
	$depletion= $_POST['depreciation'];
	$amortization=$_POST['amortization'];
	$deferredtaxes=$_POST['deferredtaxes'];
	$noncashitems=$_POST['noncashitems'];
	$changeinworkcap=$_POST['changeworkingcapital'];
	$accountreceivable=$_POST['accountreceivable'];
	$inventories=$_POST['inventories'];
	$accountspayable=$_POST['accountspayable'];
	$otherliabilities=$_POST['otherliabilities'];
	$otheropcashflow=$_POST['otheropcashflow'];
	$cashfromoperations= $_POST['cashfrmoperations'];
	
	$purchasefixassets= $_POST['purchaseoffixedassets'];
	$acquofbusiness= $_POST['acquisitbusiness'];
	$saleofbusiness= $_POST['saleofbusiness'];
	$saleoffixedassets= $_POST['saleoffixedassets'];
	$otherinvestcashflow= $_POST['otherinvestingcashflow'];
	$cashfrominvesting= $_POST['cashfrominvesting'];
	
	$fincashflow= $_POST['financingcashflow'];
	$totalcashdivpaid= $_POST['totalcashdividendspaid'];
	$issueofstock= $_POST['issuanceretirementofstock'];
	$issueofdebpt= $_POST['issuanceofdebt'];
	$cashfromfinance= $_POST['cashfromfinancing'];
	$netachangeincash= $_POST['netchangeincash'];
	$netincashbeginbal= $_POST['netcashbeginingbalance'];
	$netincashendbal= $_POST['netcashendingbalance'];
	
	$query1=
	"INSERT INTO cash_flow ( company, finyear, netincomestarting, depletion, amortization, 
	deferred_taxes, noncashitems, changeinworkcap, accountreceiv, inventories, 
	accountpayable, otherliab, otheropcashflow, cashfromoper,
	purchasefixassets, 	acqofbusiness, saleofbusiness, saleoffixedassets, otherinvestcashflow,
	cashfrominvesting, 
	fincashflow, totalcashdivpaid, issueofstock, inssueofdebt, 
	cashfromfinance, netchangeincash, netincashbeginbal, netincashendbal)
	VALUES (". $finyear . ",'" .$company. "',". $netincomestrating. "," . $depletion . "," .$amortization 
	. "," . $deferredtaxes . "," .$noncashitems . "," .$changeinworkcap . "," . $accountreceivable . ","
	. $inventories . "," . $accountpayable . "," . $otherliabilities . "," . $otheropcashflow . ","
	. $cashfromoperations . "," . $purchasefixassets . "," . $acquofbusiness . "," . $saleofbusiness 
	. "," . $saleoffixedassets . "," . $otherinvestcashflow . "," . $cashfrominvesting . "," 
	. $fincashflow . "," . $totalcashdivpaid . "," . $issueofstock . "," . $issueofdebpt . "," . $cashfromfinance
	. "," .  $netachangeincash  . "," .  $netincashbeginbal . ","  . $netincashendbal . ")";

mysqli_query($conn,$query1);
mysqli_close($conn);
header("location:cashflow.php?poststatus='record posted successfully'");
	
//	mysql_query($query1,$conn); 		
//	mysql_close($conn);

	
?>