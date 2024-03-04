<?php
$finyearbal=$_GET['finyearbal'];
$companybal=$_GET['companybal'];


$conn = mysql_connect("localhost","nacheeso_hemant","7ByV+sA[CsL6") or die(mysql_error());
mysql_select_db("nacheeso_finance");

$querybal= "DELETE FROM balance_sheet WHERE finyear=".intval($finyearbal) . " and company='" . $companybal ."'" ;
mysql_query($querybal,$conn);

header ("location:balancedelete.php");


?>
