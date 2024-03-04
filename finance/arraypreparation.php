<?php
$conn = mysqli_connect("localhost","root","","finance") or die(mysqli_error());
$query2 = "select finyear, totalrevenue_sales from income_statement where company='ABC Ltd.'";
$result = mysqli_query($conn,$query2);
$i=0;
$j=0;
$results = array();

while ($info=mysqli_fetch_array($result,MYSQLI_ASSOC))
{
	$finyear=$info['finyear'];
	$revsales=$info['totalrevenue_sales'];
	$results[$i][$j]=$finyear;
	$j=$j+1;	
	$results[$i+1][$j]=$revsales;

}
print_r($results);
?>