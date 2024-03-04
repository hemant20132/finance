<?php
$finyearcash=$_GET['finyearcash'];
$companycash=$_GET['companycash'];


$conn = mysql_connect("localhost","nacheeso_hemant","7ByV+sA[CsL6") or die(mysql_error());
mysql_select_db("nacheeso_finance");

$querycash= "DELETE FROM cash_flow WHERE finyear=".intval($finyearcash) . " and company='" . $companycash ."'" ;
mysql_query($querycash,$conn);

header ("location:cashflowdelete.php");


?>
