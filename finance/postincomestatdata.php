<?php
	ob_start();
	error_reporting(E_ALL);

	$finyear=intval($_POST['finyear']);
	$company=$_POST['company'];
	$totalrevenue_sales=intval($_POST['revsales']);
	$costrevenue=intval($_POST['costofrevenue']);
    $grossprofit=intval($_POST['grossprofit']);	
	$generaladmexpenses=intval($_POST['generaladmexpenses']);
	$researchdevelop=intval($_POST['researchdev']);
	$salesmarketingexp=intval($_POST['salesmarketingexp']);
	$depreciationamort=intval($_POST['depreciationamort']);
	$interestexp1=intval($_POST['interestexp1']);
	$businespecific1=intval($_POST['businespecific1']);
	$businespecific2=intval($_POST['businespecific2']);
	$businespecific3=intval($_POST['businespecific3']);
	$miscexp=intval($_POST['miscexp']);
	$operatingexp=intval($_POST['operatingexp']);
	$operatinginc=intval($_POST['operatinginc']);
	$interestexp2=intval($_POST['interestexp2']);
	$incomebeforetax=intval($_POST['incomebeforetax']);
	$tax=intval($_POST['tax']);
	$netprofit=intval($_POST['netprofit']);
	
	$conn=mysqli_connect('fdb3.awardspace.net','2051670_db1','mastermagic2015','2051670_db1','3306','') or  die ("Error in query:".mysqli_error());
	
//	$conn=mysql_connect("localhost","root","") or die(mysql_error());
//	mysql_select_db("finance");

	$query = "INSERT INTO 
	income_statement (finyear, company, totalrevenue_sales, costrevenue, 
	grossprofit, generaladmexpenses, researchdevelop, salesmarketingexp, depreciationamort, 
	interestexp1, businespecific1, businespecific2, businespecific3, miscexp, operatingexp, 
	operatinginc, interestexp2, incomebeforetax, tax, netprofit) 
	VALUES (" . $finyear . ",'" .$company. "'," . $totalrevenue_sales . "," . $costrevenue . "," . $grossprofit .
	"," . $generaladmexpenses . "," . $researchdevelop ."," . $salesmarketingexp . ",". $depreciationamort ."," . $interestexp1.
	"," . $businespecific1 . "," . $businespecific2 . "," . $businespecific3 . "," . $miscexp . "," . $operatingexp ."," . $operatinginc .
	"," . $interestexp2 . "," . $incomebeforetax . "," . $tax . "," . $netprofit . ")";	

	mysqli_query($conn,$query); 	
	mysqli_close($conn);
	header("location:incomestatement.php?poststatus='record posted successfully'");
		
	?>