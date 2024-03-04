<?php

$conn=mysqli_connect('fdb3.awardspace.net','2051670_db1','mastermagic2015','2051670_db1','3306','') or  die ("Error in query:".mysqli_error());
$query1="DELETE FROM income_statement WHERE 1";
mysqli_query($conn,$query1);
$query2="DELETE FROM balance_sheet WHERE 1";
mysqli_query($conn,$query2);
$query3="DELETE FROM cash_flow WHERE 1";
mysqli_query($conn,$query3);
$query4="DELETE FROM uploadfiledata WHERE 1";
mysqli_query($conn,$query4);


?>
	