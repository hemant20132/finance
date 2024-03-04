<?php
$finyearinc=$_GET['finyearinc'];
$companyinc=$_GET['companyinc'];


$conn = mysql_connect("localhost","nacheeso_hemant","7ByV+sA[CsL6") or die(mysql_error());
mysql_select_db("nacheeso_finance");

$queryinc= "DELETE FROM income_statement WHERE finyear=".intval($finyearinc) . " and company='" . $companyinc ."'" ;
mysql_query($queryinc,$conn);

header ("location:incomedelete.php");


?>
